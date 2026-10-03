import copy
import json
import tempfile
import unittest
from pathlib import Path

from jsonschema import ValidationError

from inventory import ROOT, build, make_row, scoped_slice, validate


class InventoryTests(unittest.TestCase):
    def test_provisioning_schema_routes_have_owners_in_both_versions(self):
        for version in ("v1.0", "beta"):
            for root in ("filterOperators", "functions"):
                for suffix, method, owner, disposition in [
                    ("", "GET", "EXT-03", "scheduled"),
                    ("/{id}", "GET", "EXT-03", "scheduled"),
                    ("/$count", "GET", "EXT-03", "scheduled"),
                    ("", "POST", "WRITE-N", "scheduled"),
                    ("/{id}", "PATCH", "WRITE-N", "scheduled"),
                    ("/{id}", "DELETE", "WRITE-N", "scheduled"),
                ]:
                    with self.subTest(version=version, root=root, suffix=suffix, method=method):
                        row = make_row(version, f"/{root}{suffix}", method, {"operationId": "fixture"})
                        self.assertEqual(row["owningSlice"], owner)
                        self.assertEqual(row["disposition"], "intentionally-blocked" if version == "beta" and method != "GET" else disposition)

    def test_tenant_search_is_a_read_for_each_action_representation(self):
        for action in ("tenantSearch", "microsoft.graph.tenantSearch", "microsoft.graph.managedTenants.tenantSearch"):
            with self.subTest(action=action):
                row = make_row("beta", f"/tenantRelationships/managedTenants/tenantGroups/{action}", "POST", {"operationId": "fixture"})
                self.assertEqual((row["disposition"], row["owningSlice"]), ("scheduled", "EXT-04"))
                self.assertEqual(row["permissions"]["sources"], ["https://learn.microsoft.com/en-us/graph/api/managedtenants-tenantgroup-tenantsearch?view=graph-rest-beta"])

    def test_restored_entra_families_have_dispositions_and_owners(self):
        for version in ("v1.0", "beta"):
            for path, owner in [
                ("/roleManagement/entitlementManagement/roleDefinitions", "EXT-02"),
                ("/roleManagement/entitlementManagement/roleAssignments/{id}/principal", "EXT-02"),
                ("/settings", "EXT-01"),
                ("/filteringPolicies/{id}/policyRules", "EXT-04"),
                ("/templates/deviceTemplates/{id}", "READ-10"),
                ("/admin/entra/uxSetting", "EXT-01"),
            ]:
                with self.subTest(version=version, path=path):
                    row = make_row(version, path, "GET", {"operationId": "fixture"})
                    self.assertEqual((row["disposition"], row["owningSlice"]), ("scheduled", owner))

    def test_read_actions_keep_family_ownership_in_both_versions(self):
        for version in ("v1.0", "beta"):
            for prefix, owner, actions in [
                ("/users", "READ-01", ["getAvailableExtensionProperties", "getUserOwnedObjects", "validateProperties"]),
                ("/groups", "READ-02", ["getAvailableExtensionProperties", "getUserOwnedObjects", "validateProperties", "evaluateDynamicMembership"]),
                ("/users/{user-id}", "READ-01", ["getMemberGroups", "getMemberObjects", "checkMemberGroups", "checkMemberObjects"]),
                ("/me", "READ-01", ["getMemberGroups", "getMemberObjects", "checkMemberGroups", "checkMemberObjects"]),
                ("/directoryObjects", "EXT-01", ["getByIds", "getAvailableExtensionProperties", "getUserOwnedObjects", "validateProperties"]),
                ("/groups/{group-id}", "READ-02", ["checkGrantedPermissionsForApp", "evaluateDynamicMembership"]),
                ("/identity/conditionalAccess", "READ-03", ["evaluate"]),
            ]:
                for action in actions:
                    for qualifier in ("", "microsoft.graph."):
                        path = f"{prefix}/{qualifier}{action}"
                        with self.subTest(version=version, path=path):
                            row = make_row(version, path, "POST", {"operationId": "fixture"})
                            self.assertEqual((row["disposition"], row["owningSlice"]), ("scheduled", owner))

    def test_cross_pack_descendants_are_excluded_before_owner_dispatch(self):
        for prefix in [
            "/authenticationMethodDevices/{id}/assignedTo",
            "/directory/deletedItems/{id}/graph.user",
            "/identityGovernance/entitlementManagement/accessPackageAssignments/{id}/target",
            "/invitations/invitedUser", "/me/manager",
            "/networkAccess/logs/traffic/{id}/user", "/users/{id}/directReports/{id}",
        ]:
            for version in ("v1.0", "beta"):
                for method in ("GET", "PATCH"):
                    with self.subTest(prefix=prefix, version=version, method=method):
                        row = make_row(version, f"{prefix}/mailboxSettings", method, {"operationId": "fixture"})
                        self.assertEqual((row["disposition"], row["owningSlice"]), ("excluded", None))
                        self.assertEqual(row["reason"], "Out-of-pack Mail navigation requires separate authorization.")
        for nav in ["cloudPcDevices", "managedDeviceCompliances", "windowsProtectionStates", "windowsDeviceMalwareStates", "managementTemplates"]:
            with self.subTest(nav=nav):
                row = make_row("beta", f"/tenantRelationships/managedTenants/{nav}/{{id}}", "GET", {"operationId": "fixture"})
                self.assertEqual((row["disposition"], row["owningSlice"]), ("excluded", None))
        self.assertEqual(scoped_slice("/tenantRelationships/managedTenants/conditionalAccessPolicyCoverages"), "EXT-04")
        self.assertEqual(scoped_slice("/tenantRelationships/delegatedAdminRelationships/{id}/accessAssignments"), "EXT-04")

    def test_identity_boundary_excludes_other_packs(self):
        for path, expected in [
            ("/users/{user-id}/authentication/methods", "READ-04"),
            ("/users(userPrincipalName='{userPrincipalName}')", "READ-01"),
            ("/groups/{group-id}/members", "READ-02"),
            ("/me/memberOf", "READ-01"),
            ("/users/{user-id}/messages", None),
            ("/users/{user-id}/managedDevices", None),
            ("/groups/{group-id}/drive", None),
            ("/roleManagement/deviceManagement", None),
            ("/reports/getEmailActivityCounts(period='{period}')", None),
        ]:
            with self.subTest(path=path):
                self.assertEqual(scoped_slice(path), expected)

    def test_secret_and_deprecated_routes_have_explicit_dispositions(self):
        for version, method, path, expected in [
            ("v1.0", "POST", "/applications/{application-id}/addPassword", "intentionally-blocked"),
            ("v1.0", "GET", "/directory/deviceLocalCredentials/{id}", "intentionally-blocked"),
            ("v1.0", "GET", "/informationProtection/bitlocker/recoveryKeys/{id}", "intentionally-blocked"),
            ("v1.0", "POST", "/users/{user-id}/authentication/passwordMethods/{id}/resetPassword", "intentionally-blocked"),
            ("v1.0", "GET", "/servicePrincipals/{id}/synchronization/secrets", "intentionally-blocked"),
            ("beta", "PATCH", "/users/{user-id}", "intentionally-blocked"),
            ("beta", "GET", "/identityGovernance/permissionsManagement", "deprecated"),
            ("v1.0", "GET", "/applications", "scheduled"),
        ]:
            with self.subTest(path=path):
                row = make_row(version, path, method, {"operationId": "fixture.operation"})
                self.assertEqual(row["disposition"], expected)

    def test_quoted_paths_and_both_versions_are_discovered(self):
        fixture = """openapi: 3.0.4
paths:
  /users:
    get:
      operationId: users.ListUsers
  '/users/{user-id}':
    patch:
      operationId: users.UpdateUser
  /users/{user-id}/messages:
    get:
      operationId: users.ListMessages
  /invitations/invitedUser/mailboxSettings:
    patch:
      operationId: invitations.UpdateMailboxSettings
  /unscoped:
    get:
      operationId: unscoped.List
components:
  schemas: {}
"""
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "fixture.yaml"
            source.write_text(fixture)
            inventory = build({"v1.0": source, "beta": source})
            validate(inventory)
        self.assertEqual([row["id"] for row in inventory["operations"]], [
            "beta:GET:/unscoped", "beta:GET:/users", "beta:GET:/users/{user-id}/messages",
            "beta:PATCH:/invitations/invitedUser/mailboxSettings", "beta:PATCH:/users/{user-id}",
            "v1.0:GET:/unscoped", "v1.0:GET:/users", "v1.0:GET:/users/{user-id}/messages",
            "v1.0:PATCH:/invitations/invitedUser/mailboxSettings", "v1.0:PATCH:/users/{user-id}",
        ])
        self.assertEqual(inventory["sources"][0]["excludedByRoot"], {"invitations": 1, "unscoped": 1, "users": 1})
        self.assertEqual([(row["disposition"], row["owningSlice"]) for row in inventory["operations"] if row["path"] in {"/unscoped", "/users/{user-id}/messages", "/invitations/invitedUser/mailboxSettings"}], [("excluded", None)] * 6)
        for path in ("/unscoped", "/users"):
            incomplete = copy.deepcopy(inventory)
            incomplete["operations"] = [row for row in incomplete["operations"] if row["path"] != path]
            with self.subTest(path=path), self.assertRaisesRegex(ValueError, "Missing operation"):
                validate(incomplete)
        cross_pack = copy.deepcopy(inventory)
        cross_pack["operations"][2].update(disposition="scheduled", owningSlice="READ-01")
        with self.assertRaisesRegex(ValueError, "Boundary mismatch"):
            validate(cross_pack)

    def test_excluded_operations_cannot_have_dispatch_owners(self):
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        inventory["operations"][0] = make_row("beta", "/users/{id}/mailboxSettings", "GET", {"operationId": "fixture"})
        inventory["operations"][0]["owningSlice"] = "READ-01"
        with self.assertRaises(ValidationError):
            validate(inventory)

    def test_duplicate_operation_is_rejected(self):
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        inventory["operations"].append(copy.deepcopy(inventory["operations"][0]))
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            validate(inventory)

    def test_missing_access_source_is_rejected(self):
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        inventory["operations"][0]["roles"]["sources"] = []
        with self.assertRaises(ValidationError):
            validate(inventory)

    def test_removed_operation_is_rejected(self):
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        inventory["operations"].pop()
        with self.assertRaisesRegex(ValueError, "Missing operation"):
            validate(inventory)

    def test_discovery_cannot_claim_raw_coverage(self):
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        inventory["operations"][0]["disposition"] = "reviewed-raw-read"
        with self.assertRaisesRegex(ValueError, "Discovery alone"):
            validate(inventory)


if __name__ == "__main__":
    unittest.main()
