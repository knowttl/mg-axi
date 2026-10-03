"""Offline operation discovery. This is build tooling, not the runtime catalogue."""

import argparse
import hashlib
import json
from pathlib import Path

import yaml
from jsonschema import Draft202012Validator

ROOT = Path(__file__).resolve().parents[1]
REVISION = "7b2914c8ad1340129f52aa785f13c074cb46fd7c"
METHODS = {"get", "post", "put", "patch", "delete", "head", "options"}
MAP_SOURCE = "https://learn.microsoft.com/en-us/graph/api/resources/identity-network-access-overview?view=graph-rest-1.0"
LICENCE_SOURCE = "https://learn.microsoft.com/en-us/entra/fundamentals/licensing"

# Only identity navigation properties belong to the Entra boundary.
USER_NAV = set("agreementAcceptances appConsentRequestsForApproval appRoleAssignedResources appRoleAssignments authentication createdObjects directReports extensions identities invitedBy joinedGroups licenseDetails manager memberOf oauth2PermissionGrants ownedDevices ownedObjects pendingAccessReviewInstances permissionGrants registeredDevices scopedRoleMemberOf scopedRoleMemberships serviceProvisioningErrors sponsorOf transitiveMemberOf transitiveReports onPremisesSyncBehavior sponsors authorizationInfo signInIdentifiers usageRights".split())
GROUP_NAV = set("appRoleAssignments createdOnBehalfOf endpoints extensions groupLifecyclePolicies members membersWithLicenseErrors memberOf owners permissionGrants serviceProvisioningErrors transitiveMembers transitiveMemberOf settings onPremisesSyncBehavior".split())
ROOT_SLICES = {
    "applications": "READ-07", "servicePrincipals": "READ-07",
    "directoryRoles": "READ-09", "directoryRoleTemplates": "READ-09",
    "devices": "READ-10", "directory": "EXT-01", "directoryObjects": "EXT-01",
    "directorySettings": "EXT-01", "directorySettingTemplates": "EXT-01",
    "groupSettings": "EXT-01", "groupSettingTemplates": "EXT-01",
    "organization": "EXT-01", "organizationSettings": "EXT-01",
    "domains": "EXT-01", "subscribedSkus": "EXT-01", "companySubscriptions": "EXT-01",
    "contacts": "EXT-01", "onPremisesDirectorySynchronization": "EXT-01",
    "policies": "EXT-01", "identityProviders": "EXT-03", "invitations": "EXT-03",
    "identity": "EXT-03", "identityGovernance": "EXT-02",
    "identityProtection": "READ-06", "auditLogs": "READ-05",
    "tenantRelationships": "EXT-04", "contracts": "EXT-04",
    "networkAccess": "EXT-04", "agentRegistry": "EXT-04",
    "externalUserProfiles": "EXT-03", "pendingExternalUserProfiles": "EXT-03",
    "dataPolicyOperations": "EXT-03", "certificateBasedAuthConfiguration": "EXT-01",
    "trustFramework": "EXT-03",
    "administrativeUnits": "READ-10", "applicationTemplates": "READ-07",
    "agreements": "EXT-01", "agreementAcceptances": "EXT-01",
    "schemaExtensions": "EXT-01", "domainDnsRecords": "EXT-01",
    "allowedDataLocations": "EXT-01", "groupLifecyclePolicies": "EXT-01",
    "mobilityManagementPolicies": "EXT-01", "onPremisesPublishingProfiles": "EXT-04",
    "permissionGrants": "READ-08", "scopedRoleMemberships": "READ-09",
    "authenticationMethodsPolicy": "READ-04", "authenticationMethodConfigurations": "READ-04",
    "authenticationMethodDevices": "READ-04", "certificateAuthorities": "EXT-01",
    "riskyUsers": "READ-06", "riskDetections": "READ-06",
    "accessReviews": "EXT-02", "accessReviewDecisions": "EXT-02",
    "programs": "EXT-02", "programControls": "EXT-02", "programControlTypes": "EXT-02",
    "businessFlowTemplates": "EXT-02", "governanceResources": "EXT-02",
    "governanceRoleAssignments": "EXT-02", "governanceRoleAssignmentRequests": "EXT-02",
    "governanceRoleDefinitions": "EXT-02", "governanceRoleSettings": "EXT-02",
    "governanceSubjects": "EXT-02", "privilegedRoles": "READ-09",
    "privilegedRoleAssignments": "READ-09", "privilegedRoleAssignmentRequests": "READ-09",
    "privilegedApproval": "READ-09", "privilegedOperationEvents": "READ-09",
    "privilegedSignupStatus": "READ-09",
}
USER_ACTIONS = set("assignLicense changePassword checkMemberGroups checkMemberObjects convertExternalToInternalMemberUser deletePasswordSingleSignOnCredentials exportPersonalData getMemberGroups getMemberObjects getPasswordSingleSignOnCredentials invalidateAllRefreshTokens reprocessLicenseAssignment restore retryServiceProvisioning revokeSignInSessions validatePassword".split())
GROUP_ACTIONS = set("assignLicense checkGrantedPermissionsForApp checkMemberGroups checkMemberObjects deletePasswordSingleSignOnCredentials evaluateDynamicMembership getMemberGroups getMemberObjects getPasswordSingleSignOnCredentials getByIds renew restore retryServiceProvisioning validateProperties".split())


def scoped_slice(path):
    """Return the owning family, or None for a separately authorized domain."""
    parts = path.strip("/").split("/")
    root = parts[0].split("(")[0]
    if root == "me":
        root = "users"
        parts = ["users", "{user-id}", *parts[1:]]
    if root in {"users", "groups"}:
        owner = "READ-01" if root == "users" else "READ-02"
        nav = USER_NAV if root == "users" else GROUP_NAV
        actions = USER_ACTIONS if root == "users" else GROUP_ACTIONS
        tail = parts[1:]
        if tail and (tail[0].startswith("{") or tail[0] == "microsoft.graph.user" or tail[0] == "microsoft.graph.group"):
            tail = tail[1:]
        if tail:
            first = tail[0].removeprefix("microsoft.graph.").split("(")[0]
            if first not in nav | actions | {"$count", "delta", "getByIds"}:
                return None
            if first == "authentication":
                return "READ-04"
            if first in {"appRoleAssignments", "oauth2PermissionGrants"}:
                return "READ-08"
        return owner
    if root == "roleManagement":
        return "READ-09" if len(parts) == 1 or parts[1] == "directory" else None
    if root == "privilegedAccess":
        return "READ-09" if len(parts) == 1 or parts[1].startswith(("aad", "microsoft.graph.aad")) else None
    if root == "oauth2PermissionGrants":
        return "READ-08"
    if root == "reports":
        if len(parts) > 1 and parts[1].startswith("authenticationMethods"):
            return "READ-04"
        if len(parts) > 1 and (parts[1] in {"azureADPremiumLicenseInsight", "correlations", "healthMonitoring", "identityAnalytics", "sla"} or any(word in parts[1].lower() for word in ("credential", "registration", "application", "serviceprincipal", "agent", "directory", "authentication", "signin", "tenant", "permission", "relyingparty"))):
            return "EXT-04"
        return None
    if root == "informationProtection":
        return "READ-10" if len(parts) > 1 and parts[1] == "bitlocker" else None
    if root == "identity" and len(parts) > 1 and parts[1] == "conditionalAccess":
        return "READ-03"
    if root == "directory" and len(parts) > 1 and parts[1] == "administrativeUnits":
        return "READ-10"
    if root in {"applications", "servicePrincipals"} and "synchronization" in parts:
        return "EXT-03"
    if root == "servicePrincipals" and "appRoleAssignments" in parts:
        return "READ-08"
    return ROOT_SLICES.get(root)


def discover(source):
    """Parse one path at a time rather than retaining Graph's response schemas."""
    block = []
    in_paths = False
    with source.open() as stream:
        for line in stream:
            if line == "paths:\n":
                in_paths = True
                continue
            if not in_paths:
                continue
            if not line.startswith(" ") and line.strip():
                break
            if line.startswith("  ") and len(line) > 2 and line[2] != " " and block:
                yield from parse_path(block)
                block = []
            block.append(line)
    if block:
        yield from parse_path(block)


def parse_path(block):
    path, item = next(iter(yaml.load("".join(block), Loader=yaml.CSafeLoader).items()))
    for method, operation in item.items():
        if method in METHODS:
            yield path, method.upper(), operation


def make_row(version, path, method, operation):
    owner = scoped_slice(path)
    if owner is None:
        return None
    disposition, reason = "scheduled", "No implemented command or reviewed raw contract yet."
    lower = path.lower()
    secret = any(word in lower for word in ("addpassword", "passwordprofile", "getpassword", "resetpassword", "generatepassword", "uploadsecret", "getsecret", "/secrets", "devicelocalcredentials/", "bitlocker/recoverykeys/")) and not lower.endswith("/$count")
    secret = secret or (method == "POST" and "temporaryaccesspassmethods" in lower)
    if operation.get("deprecated") or "permissionsmanagement" in lower:
        disposition, reason = "deprecated", "Deprecated metadata or multicloud permissions management."
    elif secret or "trustframework/keysets" in lower:
        disposition, reason = "intentionally-blocked", "Credential values, recovery keys, LAPS passwords or secret minting."
    elif version == "beta" and method != "GET":
        disposition, reason = "intentionally-blocked", "Beta writes are denied by the approved plan."
    elif any(word in lower for word in ("b2c", "authenticationeventsflows", "trustframework/policies")):
        disposition, reason = "intentionally-blocked", "External-customer launch support requires separate authorization."
    if method != "GET":
        owner = "WRITE-N"
        if method == "PATCH" and path == "/users/{user-id}":
            owner = "WRITE-02"
        elif method == "POST" and path == "/groups/{group-id}/members/$ref":
            owner = "WRITE-01"
        elif path.endswith(("/microsoft.graph.revokeSignInSessions", "/revokeSignInSessions")):
            owner = "WRITE-03"
        elif method == "PATCH" and path == "/identity/conditionalAccess/policies/{conditionalAccessPolicy-id}":
            owner = "WRITE-04"
        elif path.endswith(("/riskyUsers/microsoft.graph.dismiss", "/riskyUsers/dismiss")):
            owner = "WRITE-05"
    doc = operation.get("externalDocs", {}).get("url")
    source = doc or MAP_SOURCE
    return {
        "id": f"{version}:{method}:{path}", "operationId": operation["operationId"],
        "version": version, "method": method, "path": path, "cloud": "commercial",
        "tenantType": "workforce", "documentation": doc,
        "authModes": {mode: {"status": "not-reviewed", "sources": [source]} for mode in ("delegated", "application")},
        "permissions": {"status": "not-reviewed", "sources": [source]},
        "roles": {"status": "not-reviewed", "sources": [source]},
        "licences": {"status": "not-reviewed", "sources": [source, LICENCE_SOURCE]},
        "cloudAvailability": {"status": "not-reviewed", "sources": [source]},
        "disposition": disposition, "reason": reason, "owningSlice": owner,
    }


def build(sources):
    rows, pins = [], []
    for version, source in sources.items():
        raw = source.read_bytes()
        pins.append({"version": version, "path": f"openapi/{version}/openapi.yaml", "sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)})
        discovered, excluded = 0, {}
        for path, method, operation in discover(source):
            discovered += 1
            row = make_row(version, path, method, operation)
            if row:
                rows.append(row)
            else:
                root = path.split("/")[1].split("(")[0]
                excluded[root] = excluded.get(root, 0) + 1
        pins[-1]["discoveredOperations"] = discovered
        pins[-1]["excludedByRoot"] = dict(sorted(excluded.items()))
    rows.sort(key=lambda row: row["id"])
    return {"schemaVersion": 1, "repository": "microsoftgraph/msgraph-metadata", "revision": REVISION, "checkedOn": "2026-10-03", "sources": pins, "operations": rows}


def validate(inventory):
    schema = json.loads((ROOT / "inventory/schema.json").read_text())
    Draft202012Validator(schema).validate(inventory)
    ids = set()
    for row in inventory["operations"]:
        if row["id"] in ids or row["id"] != f'{row["version"]}:{row["method"]}:{row["path"]}':
            raise ValueError("Duplicate or inconsistent operation identity")
        ids.add(row["id"])
        if row["disposition"] in {"named-command", "reviewed-raw-read"}:
            raise ValueError("Discovery alone cannot claim implementation or reviewed raw access")
    if {source["version"] for source in inventory["sources"]} != {"v1.0", "beta"}:
        raise ValueError("Both metadata versions must be pinned independently")
    for source in inventory["sources"]:
        scoped_count = sum(row["version"] == source["version"] for row in inventory["operations"])
        if scoped_count + sum(source["excludedByRoot"].values()) != source["discoveredOperations"]:
            raise ValueError("Missing operation: scoped and excluded counts must reconcile")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("generate", "check", "validate"))
    parser.add_argument("--source-dir", type=Path, default=ROOT / ".cache/inventory")
    args = parser.parse_args()
    output = ROOT / "inventory/operations.json"
    if args.action == "validate":
        inventory = json.loads(output.read_text())
        validate(inventory)
        print(f'Validated {len(inventory["operations"])} discovered operation rows; no completeness percentage claimed.')
        return
    inventory = build({version: args.source_dir / f"{version}.yaml" for version in ("v1.0", "beta")})
    validate(inventory)
    # One operation per line keeps generated diffs localized to changed routes.
    header = {key: value for key, value in inventory.items() if key != "operations"}
    encoded = json.dumps(header, indent=2)[:-2] + ',\n  "operations": [\n'
    encoded += ",\n".join("    " + json.dumps(row) for row in inventory["operations"])
    encoded += "\n  ]\n}\n"
    if args.action == "generate":
        output.parent.mkdir(exist_ok=True)
        output.write_text(encoded)
    elif output.read_text() != encoded:
        raise ValueError("Inventory drift: regenerate and review the metadata/boundary diff")
    print(f'Validated {len(inventory["operations"])} discovered operation rows; no completeness percentage claimed.')


if __name__ == "__main__":
    main()
