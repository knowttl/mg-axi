import copy
import json
import tempfile
import unittest
from pathlib import Path

from jsonschema import ValidationError

from inventory import ROOT, build, make_row, scoped_slice, validate


class InventoryTests(unittest.TestCase):
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
