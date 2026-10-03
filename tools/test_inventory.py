import copy
import json
import tempfile
import unittest
from pathlib import Path

from jsonschema import ValidationError

from inventory import ROOT, build, make_row, scoped_slice, validate


class InventoryTests(unittest.TestCase):
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
                        self.assertIsNone(make_row(version, f"{prefix}/mailboxSettings", method, {"operationId": "fixture"}))
        for nav in ["cloudPcDevices", "managedDeviceCompliances", "windowsProtectionStates", "windowsDeviceMalwareStates", "managementTemplates"]:
            with self.subTest(nav=nav):
                self.assertIsNone(scoped_slice(f"/tenantRelationships/managedTenants/{nav}/{{id}}"))
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
components:
  schemas: {}
"""
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "fixture.yaml"
            source.write_text(fixture)
            inventory = build({"v1.0": source, "beta": source})
            validate(inventory)
        self.assertEqual([row["id"] for row in inventory["operations"]], [
            "beta:GET:/users", "beta:PATCH:/users/{user-id}",
            "v1.0:GET:/users", "v1.0:PATCH:/users/{user-id}",
        ])
        self.assertEqual(inventory["sources"][0]["excludedByRoot"], {"users": 1})

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
