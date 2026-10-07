import copy
import json
import tempfile
import unittest
from pathlib import Path

from jsonschema import ValidationError

from inventory import CUTOFF_REASON, ROOT, build, make_row, scoped_slice, validate


class InventoryTests(unittest.TestCase):
    def test_identity_provider_reads_defer_only_beta(self):
        for suffix in ("", "/{identityProviderBase-id}", "/$count", "/availableProviderTypes()"):
            for version, reason in [
                ("v1.0", "No implemented command or reviewed raw contract yet."),
                ("beta", "Deferred to a later EXT-03 beta identity-providers subfamily: Identity-provider reads support v1.0 only; beta needs its own review."),
            ]:
                with self.subTest(suffix=suffix, version=version):
                    row = make_row(version, f"/identity/identityProviders{suffix}", "GET", {"operationId": "fixture"})
                    self.assertEqual((row["disposition"], row["owningSlice"], row["reason"]), ("scheduled", "EXT-03", reason))

    def test_invitation_reads_are_unavailable_without_a_documented_get_contract(self):
        for suffix in ("", "/$count", "/invitedUser", "/invitedUser/serviceProvisioningErrors", "/invitedUser/serviceProvisioningErrors/$count", "/invitedUserSponsors", "/invitedUserSponsors/$count", "/invitedUserSponsors/{directoryObject-id}"):
            with self.subTest(suffix=suffix):
                row = make_row("v1.0", f"/invitations{suffix}", "GET", {"operationId": "fixture"})
                self.assertEqual((row["disposition"], row["owningSlice"]), ("unavailable", "EXT-03"))
                self.assertTrue(row["reason"].startswith("Marked unavailable by firstmate EXT-03b decision: "))
        with self.subTest(suffix="beta stays scheduled"):
            row = make_row("beta", "/invitations", "GET", {"operationId": "fixture"})
            self.assertEqual((row["disposition"], row["owningSlice"], row["reason"]), ("scheduled", "EXT-03", "No implemented command or reviewed raw contract yet."))
        with self.subTest(suffix="mailboxSettings stays excluded"):
            row = make_row("v1.0", "/invitations/invitedUser/mailboxSettings", "GET", {"operationId": "fixture"})
            self.assertEqual((row["disposition"], row["owningSlice"]), ("excluded", None))
        with self.subTest(suffix="create stays a write"):
            row = make_row("v1.0", "/invitations", "POST", {"operationId": "fixture"})
            self.assertEqual((row["disposition"], row["owningSlice"]), ("scheduled", "WRITE-N"))

    def test_workflow_task_validation_is_a_read_across_resources_and_names(self):
        for version in ("v1.0", "beta"):
            for resource in ("workflows", "deletedItems/workflows"):
                for qualifier in ("", "microsoft.graph.", "microsoft.graph.identityGovernance."):
                    for deprecated, disposition in [(False, "scheduled"), (True, "deprecated")]:
                        path = f"/identityGovernance/lifecycleWorkflows/{resource}/{{workflow-id}}/{qualifier}previewTaskFailures"
                        with self.subTest(version=version, path=path, deprecated=deprecated):
                            row = make_row(version, path, "POST", {"operationId": "fixture", "deprecated": deprecated})
                            self.assertEqual((row["disposition"], row["owningSlice"]), (disposition, "EXT-02"))
                            self.assertEqual(row["permissions"]["sources"], [f"https://learn.microsoft.com/en-us/graph/api/identitygovernance-workflow-previewtaskfailures?view=graph-rest-{'beta' if version == 'beta' else '1.0'}"])

    def test_confirmed_read_siblings_keep_owners_and_sources(self):
        for version in ("v1.0", "beta"):
            for prefix, action, namespace, owner, source in [
                ("/identityGovernance/entitlementManagement/accessPackages/{id}", "getApplicablePolicyRequirements", "identityGovernance", "EXT-02", "accesspackage-getapplicablepolicyrequirements"),
                ("/identityGovernance/entitlementManagement/accessPackages(uniqueName='{uniqueName}')", "getApplicablePolicyRequirements", "identityGovernance", "EXT-02", "accesspackage-getapplicablepolicyrequirements"),
                ("/identity/customAuthenticationExtensions", "validateAuthenticationConfiguration", "", "EXT-03", "customauthenticationextension-validateauthenticationconfiguration"),
                ("/identity/customAuthenticationExtensions/{id}", "validateAuthenticationConfiguration", "", "EXT-03", "customauthenticationextension-validateauthenticationconfiguration"),
                ("/applications/{id}/synchronization/jobs", "validateCredentials", "", "EXT-03", "synchronization-synchronizationjob-validatecredentials"),
                ("/applications/{id}/synchronization/jobs/{job-id}", "validateCredentials", "", "EXT-03", "synchronization-synchronizationjob-validatecredentials"),
                ("/servicePrincipals/{id}/synchronization/jobs", "validateCredentials", "", "EXT-03", "synchronization-synchronizationjob-validatecredentials"),
                ("/servicePrincipals/{id}/synchronization/jobs/{job-id}", "validateCredentials", "", "EXT-03", "synchronization-synchronizationjob-validatecredentials"),
                ("/users", "validatePassword", "", "READ-01", "user-validatepassword"),
            ]:
                for qualifier in ("", "microsoft.graph." + (namespace + "." if namespace else "")):
                    path = f"{prefix}/{qualifier}{action}"
                    with self.subTest(version=version, path=path):
                        row = make_row(version, path, "POST", {"operationId": "fixture"})
                        self.assertEqual((row["disposition"], row["owningSlice"]), ("scheduled", owner))
                        self.assertEqual(row["permissions"]["sources"], [f"https://learn.microsoft.com/en-us/graph/api/{source}?view=graph-rest-{'beta' if version == 'beta' else '1.0'}"])

    def test_read_effects_preserve_write_secret_and_exclusion_policies(self):
        for version, path, deprecated, disposition, owner in [
            ("v1.0", "/identityGovernance/lifecycleWorkflows/workflows/{id}/microsoft.graph.identityGovernance.previewWorkflow", False, "scheduled", "WRITE-N"),
            ("beta", "/identityGovernance/lifecycleWorkflows/workflows/{id}/microsoft.graph.identityGovernance.previewWorkflow", False, "intentionally-blocked", "WRITE-N"),
            ("beta", "/identityGovernance/lifecycleWorkflows/deletedItems/workflows/{id}/microsoft.graph.identityGovernance.previewWorkflow", True, "deprecated", "WRITE-N"),
            ("beta", "/identityGovernance/permissionsAnalytics/aws/findings", False, "intentionally-blocked", "WRITE-N"),
            ("beta", "/identityGovernance/lifecycleWorkflows/workflows/{id}/microsoft.graph.identityGovernance.previewUnknown", False, "intentionally-blocked", "WRITE-N"),
            ("v1.0", "/users/{id}/getPasswordSingleSignOnCredentials", False, "intentionally-blocked", "READ-01"),
            ("beta", "/servicePrincipals/{id}/getPasswordSingleSignOnCredentials", False, "intentionally-blocked", "READ-07"),
            ("beta", "/users/{id}/mailboxSettings/microsoft.graph.identityGovernance.previewTaskFailures", False, "excluded", None),
        ]:
            with self.subTest(version=version, path=path):
                row = make_row(version, path, "POST", {"operationId": "fixture", "deprecated": deprecated})
                self.assertEqual((row["disposition"], row["owningSlice"]), (disposition, owner))

    def test_generated_read_action_rows_keep_read_owners(self):
        actions = set("checkGrantedPermissionsForApp checkMemberGroups checkMemberObjects evaluate evaluateDynamicMembership getApplicablePolicyRequirements getAvailableExtensionProperties getByIds getMemberGroups getMemberObjects getUserOwnedObjects previewTaskFailures tenantSearch validateAuthenticationConfiguration validateCredentials validatePassword validateProperties".split())
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        observed = set()
        for row in inventory["operations"]:
            action = row["path"].rsplit("/", 1)[-1].rsplit(".", 1)[-1]
            if row["method"] == "POST" and row["disposition"] != "excluded" and action in actions:
                with self.subTest(id=row["id"]):
                    self.assertFalse(row["owningSlice"].startswith("WRITE-"))
                    self.assertIn(row["disposition"], {"scheduled", "deferred", "deprecated"})
                observed.add(action)
        self.assertEqual(observed, actions)

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
            inventory = build({"v1.0": source, "beta": source}, shipped={}, reviewed={}, alternates={})
            validate(inventory, shipped={}, reviewed={}, alternates={})
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

    def test_shipped_v1_reads_carry_named_command(self):
        shipped = {"GET:/users": "entra user list"}
        row = make_row("v1.0", "/users", "GET", {"operationId": "fixture"}, shipped)
        self.assertEqual((row["disposition"], row["owningSlice"], row["reason"]), ("named-command", "READ-01", "Shipped as `mg-axi entra user list`."))
        with self.subTest(reason="beta never joins the shipped v1.0 command"):
            beta = make_row("beta", "/users", "GET", {"operationId": "fixture"}, shipped)
            self.assertEqual(beta["disposition"], "scheduled")
        with self.subTest(reason="no shipped set means pure discovery"):
            plain = make_row("v1.0", "/users", "GET", {"operationId": "fixture"})
            self.assertEqual(plain["disposition"], "scheduled")

    def test_unshipped_v1_reads_defer_out_of_v1(self):
        row = make_row("v1.0", "/groups", "GET", {"operationId": "fixture"}, {})
        self.assertEqual((row["disposition"], row["owningSlice"], row["reason"]), ("deferred", "READ-02", CUTOFF_REASON))
        with self.subTest(reason="a firstmate deferral reason survives the flip"):
            deferred = make_row("v1.0", "/organization/{organization-id}/branding/favicon", "GET", {"operationId": "fixture"}, {})
            self.assertEqual(deferred["disposition"], "deferred")
            self.assertTrue(deferred["reason"].startswith("Deferred by firstmate organization scope "))
        with self.subTest(reason="a split note yields to the cutoff reason"):
            split = make_row("v1.0", "/contacts/{orgContact-id}/memberOf", "GET", {"operationId": "fixture"}, {})
            self.assertEqual((split["disposition"], split["reason"]), ("deferred", CUTOFF_REASON))

    @staticmethod
    def _inventory(rows):
        def pin(version):
            subset = [row for row in rows if row["version"] == version]
            excluded = {}
            for row in subset:
                if row["disposition"] == "excluded":
                    root = row["path"].split("/")[1].split("(")[0]
                    excluded[root] = excluded.get(root, 0) + 1
            return {"version": version, "path": f"openapi/{version}/openapi.yaml", "sha256": "0" * 64,
                    "bytes": 1, "discoveredOperations": len(subset), "excludedByRoot": excluded}
        return {"schemaVersion": 2, "repository": "microsoftgraph/msgraph-metadata", "revision": "0" * 40,
                "checkedOn": "2026-10-03", "sources": [pin("v1.0"), pin("beta")], "operations": rows}

    def test_validate_accepts_catalogue_backed_named_commands(self):
        rows = [make_row("v1.0", "/users", "GET", {"operationId": "a"}, {"GET:/users": "entra user list"}),
                make_row("beta", "/groups", "GET", {"operationId": "b"}, {})]
        validate(self._inventory(rows), shipped={"GET:/users": "entra user list"}, reviewed={}, alternates={})

    def test_reviewed_raw_routes_flip_to_reviewed_raw_read(self):
        row = make_row("v1.0", "/contacts/{orgContact-id}/transitiveMemberOf", "GET", {"operationId": "fixture"}, {}, {"v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf"})
        self.assertEqual((row["disposition"], row["owningSlice"]), ("reviewed-raw-read", "EXT-01"))
        with self.subTest(reason="a named command wins over a reviewed raw route"):
            both = make_row("v1.0", "/users", "GET", {"operationId": "a"}, {"GET:/users": "entra user list"}, {"v1.0:GET:/users"})
            self.assertEqual(both["disposition"], "named-command")

    def test_validate_rejects_unbacked_implementation_claims(self):
        rows = [make_row("v1.0", "/users", "GET", {"operationId": "a"}, {}),
                make_row("beta", "/groups", "GET", {"operationId": "b"}, {})]
        rows[0]["disposition"] = "named-command"
        with self.assertRaisesRegex(ValueError, "Discovery alone"):
            validate(self._inventory(rows), shipped={}, reviewed={}, alternates={})

    def test_validate_rejects_shipped_operations_still_scheduled(self):
        rows = [make_row("v1.0", "/users", "GET", {"operationId": "a"}),
                make_row("beta", "/groups", "GET", {"operationId": "b"})]
        with self.assertRaisesRegex(ValueError, "not named-command"):
            validate(self._inventory(rows), shipped={"GET:/users": "entra user list"}, reviewed={}, alternates={})
        with self.subTest(reason="a shipped operation left deferred fails"):
            rows[0]["disposition"] = "deferred"
            with self.assertRaisesRegex(ValueError, "not named-command"):
                validate(self._inventory(rows), shipped={"GET:/users": "entra user list"}, reviewed={}, alternates={})

    def test_validate_accepts_review_backed_raw_reads(self):
        rows = [make_row("v1.0", "/contacts/{orgContact-id}/transitiveMemberOf", "GET", {"operationId": "a"}, {}, {"v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf"}),
                make_row("beta", "/groups", "GET", {"operationId": "b"}, {})]
        validate(self._inventory(rows), shipped={}, reviewed={"v1.0:GET:/contacts/{orgContact-id}/transitiveMemberOf"}, alternates={})

    def test_validate_rejects_reviewed_routes_still_scheduled(self):
        rows = [make_row("v1.0", "/users", "GET", {"operationId": "a"}),
                make_row("beta", "/groups", "GET", {"operationId": "b"})]
        with self.assertRaisesRegex(ValueError, "not reviewed-raw-read"):
            validate(self._inventory(rows), shipped={}, reviewed={"v1.0:GET:/users"}, alternates={})
        with self.subTest(reason="a reviewed route left deferred fails"):
            rows[0]["disposition"] = "deferred"
            with self.assertRaisesRegex(ValueError, "not reviewed-raw-read"):
                validate(self._inventory(rows), shipped={}, reviewed={"v1.0:GET:/users"}, alternates={})

    def test_table_alternates_flip_to_named_command(self):
        alternates = {"GET:/groups/{group-id}/transitiveMembers": "entra group member list"}
        row = make_row("v1.0", "/groups/{group-id}/transitiveMembers", "GET", {"operationId": "fixture"}, {}, set(), alternates)
        self.assertEqual((row["disposition"], row["owningSlice"], row["reason"]),
                         ("named-command", "READ-02", "Shipped as `mg-axi entra group member list` alternate route."))
        with self.subTest(reason="a named command wins over a table alternate"):
            both = make_row("v1.0", "/users", "GET", {"operationId": "a"}, {"GET:/users": "entra user list"}, set(), {"GET:/users": "entra user list"})
            self.assertEqual(both["reason"], "Shipped as `mg-axi entra user list`.")
        with self.subTest(reason="a table alternate wins over a reviewed raw route"):
            overlap = make_row("v1.0", "/groups/{group-id}/transitiveMembers", "GET", {"operationId": "fixture"}, {}, {"v1.0:GET:/groups/{group-id}/transitiveMembers"}, alternates)
            self.assertEqual(overlap["disposition"], "named-command")

    def test_validate_accepts_table_backed_alternates(self):
        alternates = {"GET:/groups/{group-id}/transitiveMembers": "entra group member list"}
        rows = [make_row("v1.0", "/groups/{group-id}/transitiveMembers", "GET", {"operationId": "a"}, {}, set(), alternates),
                make_row("beta", "/groups", "GET", {"operationId": "b"}, {})]
        validate(self._inventory(rows), shipped={}, reviewed={}, alternates=alternates)
        with self.subTest(reason="a table-backed alternate still scheduled fails"):
            rows[0]["disposition"] = "scheduled"
            with self.assertRaisesRegex(ValueError, "not named-command"):
                validate(self._inventory(rows), shipped={}, reviewed={}, alternates=alternates)
        with self.subTest(reason="a table-backed alternate left deferred fails"):
            rows[0]["disposition"] = "deferred"
            with self.assertRaisesRegex(ValueError, "not named-command"):
                validate(self._inventory(rows), shipped={}, reviewed={}, alternates=alternates)

    def test_alternate_without_a_shipped_command_fails(self):
        from inventory import alternate_operations
        with self.assertRaisesRegex(ValueError, "no shipped command behind it"):
            alternate_operations({})
        with self.subTest(reason="attribution is deterministic"):
            from inventory import shipped_operations
            first = alternate_operations(shipped_operations())
            self.assertEqual(first, alternate_operations(shipped_operations()))
            self.assertEqual(len(first), 23)

    def test_discovery_cannot_claim_raw_coverage(self):
        inventory = json.loads((ROOT / "inventory/operations.json").read_text())
        inventory["operations"][0]["disposition"] = "reviewed-raw-read"
        with self.assertRaisesRegex(ValueError, "Discovery alone"):
            validate(inventory)


if __name__ == "__main__":
    unittest.main()
