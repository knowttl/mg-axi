import { encode } from "@toon-format/toon";
import { AxiError, runAxiCli } from "axi-sdk-js";
import { LEAVES, home, leafHelp, operationFor, resolveCommand, DESCRIPTION, TOP_LEVEL_HELP } from "./catalogue.js";
import { VERSION } from "./version.js";
import { Profiles } from "./profiles.js";
import { GraphSession, MAX_CURSOR_BYTES, type GraphTransport } from "./graph-session.js";
import { listUsers, showUser } from "./entra-users.js";
import { updateUserAccount } from "./entra-user-update.js";
import { revokeUserSessions } from "./entra-user-revoke-sessions.js";
import { createMutationCoordinator, mutationFetchTransport, type MutationTransport } from "./mutations.js";
import { TRANSITIVE_OPERATION, listGroupMemberOf, listGroupMembers, listGroups, showGroup } from "./entra-groups.js";
import { listDirectoryRoles, showDirectoryRole, listRoleAssignments, listPimEligible, listPimActive } from "./entra-roles.js";
import { listAdministrativeUnitMembers, listAdministrativeUnits, listDevices, showAdministrativeUnit, showDevice } from "./entra-directory.js";
import { listSignIns, showSignIn, listDirectoryAudits, showDirectoryAudit } from "./entra-audit-logs.js";
import { listApplicationOwners, listApplications, listServicePrincipalOwners, listServicePrincipals, showApplication, showServicePrincipal } from "./entra-apps.js";
import { listAppRoleAssignments, listOAuth2Grants } from "./entra-grants.js";
import { listRiskyUsers, showRiskyUser, listRiskDetections, showRiskDetection } from "./entra-risk.js";
import { dismissRiskyUser } from "./entra-risk-dismiss.js";
import { listPolicies, showPolicy, listNamedLocations, showNamedLocation } from "./entra-conditional-access.js";
import { listBrandingLocalizations, listOrganizations, showBranding, showBrandingLocalization, showOrganization } from "./entra-organization.js";
import { updateCaPolicy } from "./entra-ca-policy-update.js";
import { listDomains, showDomain, listVerificationDnsRecords, showVerificationDnsRecord, listServiceConfigurationRecords, showServiceConfigurationRecord, listDomainDnsRecords, showDomainDnsRecord } from "./entra-domains.js";
import { listCertificateAuthPkis, showCertificateAuthPki, countCertificateAuthPkis, listCertificateAuthorities, showCertificateAuthority, countCertificateAuthorities } from "./entra-certificate-auth.js";
import { listSubscriptions, showSubscription, countSubscriptions } from "./entra-subscriptions.js";
import { listSynchronizations, showSynchronization } from "./entra-on-premises-synchronization.js";
import { listAgreements, showAgreement, listAgreementAcceptances, showAgreementAcceptance, listAcceptances, showAcceptance } from "./entra-terms-of-use.js";
import { countContracts, listContracts, showContract } from "./entra-contracts.js";
import { countLifecyclePolicies, listLifecyclePolicies, showLifecyclePolicy, countSettingTemplates, listSettingTemplates, showSettingTemplate } from "./entra-group-lifecycle.js";
import { listAttributeSets, showAttributeSet, countAttributeSets, listCustomSecurityAttributeDefinitions, showCustomSecurityAttributeDefinition, countCustomSecurityAttributeDefinitions, listAllowedValues, showAllowedValue, countAllowedValues } from "./entra-custom-security-attributes.js";
import { listIdentityProviders, showIdentityProvider, countIdentityProviders, availableIdentityProviderTypes } from "./entra-identity-providers.js";
import { listDefinitions, showDefinition, listInstances, showInstance, listDecisions, showDecision, listContactedReviewers, showContactedReviewer, listStages, showStage } from "./entra-access-reviews.js";
import { listAuthenticationMethods, listRegistrationDetails } from "./entra-auth-methods.js";
import { addGroupMember } from "./entra-group-member-add.js";
import { fetchTransport } from "./api.js";
import { doctorTargets, runDoctor } from "./doctor.js";
import { setupView } from "./setup.js";
import type { DelegatedAuth } from "./auth.js";
import type { ApplicationAuth } from "./app-auth.js";

function localHome(name?: string) {
  const output = home();
  const store = new Profiles();
  if (name || store.list().length) {
    try {
      const selected = store.resolve(name);
      output.profile = selected.name;
      output.tenant = selected.profile.tenantId;
    } catch (error) {
      if (name || !(error instanceof AxiError) || error.code !== "AUTH_REQUIRED") throw error;
      output.profile = "unavailable: no default profile selected";
    }
  }
  output.help.unshift("mg-axi profile list", "mg-axi login --help");
  return output;
}

// Test seam over true external boundaries only: the packaged dispatch stays
// identical while offline journeys substitute fixture credential services and
// transports. Production callers pass no overrides and reach MSAL + HTTPS.
export interface DispatchOverrides {
  transport?: GraphTransport;
  mutationTransport?: MutationTransport;
  delegated?: DelegatedAuth;
  application?: ApplicationAuth;
  journalPath?: string;
}

async function readCursor(cursor: string | undefined): Promise<string | undefined> {
  if (cursor !== "-") return cursor;
  const chunks: Buffer[] = [];
  let bytes = 0;
  for await (const chunk of process.stdin) {
    const buffer = Buffer.from(chunk);
    bytes += buffer.length;
    if (bytes > MAX_CURSOR_BYTES) throw new AxiError(`Collection cursor exceeds ${MAX_CURSOR_BYTES} bytes`, "VALIDATION_ERROR", ["Use a cursor within the supported size ceiling"]);
    chunks.push(buffer);
  }
  return Buffer.concat(chunks).toString("utf8");
}

export async function executeArgv(argv: string[], overrides: DispatchOverrides = {}): Promise<string | Record<string, unknown>> {
  const { leaf, flags, positional } = resolveCommand(argv);
  if (flags.help) return leafHelp(leaf);
  if (leaf.path === "home") return localHome(flags.profile as string | undefined);
  const profiles = new Profiles();
  if (leaf.path === "profile create") {
    if (String(flags.mode ?? "delegated") !== "application" && (flags["certificate-thumbprint"] !== undefined || flags.federated)) throw new AxiError("Certificate and federated credentials belong to application profiles; pass --mode application", "VALIDATION_ERROR", [leafHelp(leaf)]);
    return profiles.create(String(flags.name), String(flags.tenant), String(flags.client), String(flags.cloud), !!flags["allow-device-code"],
    String(flags.mode ?? "delegated") === "application"
      ? { certificateThumbprint: flags["certificate-thumbprint"] === undefined ? undefined : String(flags["certificate-thumbprint"]), federated: !!flags.federated }
      : undefined);
  }
  if (leaf.path === "profile list") {
    const items = profiles.list();
    return items.length ? { profiles: items, help: ["mg-axi profile show --profile <name>", "mg-axi login --help"] } : { profiles: "0 profiles configured", help: ["mg-axi profile create --help"] };
  }
  if (leaf.path === "profile show") return profiles.resolve(flags.profile as string | undefined);
  if (leaf.path === "setup") return setupView(profiles);
  if (leaf.path === "doctor") {
    const names = doctorTargets(profiles, flags.profile as string | undefined);
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    // A failing profile still reports its bounded read on stdout; the
    // nonzero exit travels through process state, never the payload.
    const result = await runDoctor({ store: profiles, names, session });
    if (result.failed) process.exitCode = 1;
    return result.output;
  }
  if (leaf.path === "login") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    if (selected.profile.mode !== "delegated") throw new AxiError("Application profiles authenticate with client credentials; interactive login is unavailable", "VALIDATION_ERROR", ["mg-axi profile show --profile <name>", "Application tokens are acquired silently with the configured Graph .default audience"]);
    const { DelegatedAuth } = await import("./auth.js");
    const { MsalProvider } = await import("./msal-provider.js");
    return { profile: selected.name, ...await new DelegatedAuth(new MsalProvider()).login(selected.profile, String(flags.method ?? "browser"), String(flags.scopes).split(",")) };
  }
  if (leaf.path === "entra conditional-access policy list" || leaf.path === "entra conditional-access policy show" || leaf.path === "entra conditional-access named-location list" || leaf.path === "entra conditional-access named-location show") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    if (flags.cursor !== undefined) flags.cursor = (await readCursor(String(flags.cursor)))!;
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra conditional-access policy list": return listPolicies(session, flags, selected.profile, operation, help, selected.name);
      case "entra conditional-access policy show": return showPolicy(session, flags, selected.profile, operation, help, selected.name);
      case "entra conditional-access named-location list": return listNamedLocations(session, flags, selected.profile, operation, help, selected.name);
      default: return showNamedLocation(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra access-review definition list" || leaf.path === "entra access-review definition show" || leaf.path === "entra access-review instance list" || leaf.path === "entra access-review instance show" || leaf.path === "entra access-review decision list" || leaf.path === "entra access-review decision show" || leaf.path === "entra access-review contacted-reviewer list" || leaf.path === "entra access-review contacted-reviewer show" || leaf.path === "entra access-review stage list" || leaf.path === "entra access-review stage show") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    if (flags.cursor !== undefined) flags.cursor = (await readCursor(String(flags.cursor)))!;
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra access-review definition list": return listDefinitions(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review definition show": return showDefinition(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review instance list": return listInstances(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review instance show": return showInstance(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review decision show": return showDecision(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review contacted-reviewer list": return listContactedReviewers(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review contacted-reviewer show": return showContactedReviewer(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review stage list": return listStages(session, flags, selected.profile, operation, help, selected.name);
      case "entra access-review stage show": return showStage(session, flags, selected.profile, operation, help, selected.name);
      default: return listDecisions(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "api get") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const { runApiGet } = await import("./api.js");
    const { DelegatedAuth } = await import("./auth.js");
    const { ApplicationAuth } = await import("./app-auth.js");
    const { MsalProvider } = await import("./msal-provider.js");
    const { MsalApplicationProvider } = await import("./msal-app-provider.js");
    const cursor = await readCursor(flags.cursor === undefined ? undefined : String(flags.cursor));
    return runApiGet({
      path: positional!,
      apiVersion: String(flags["api-version"] ?? "v1.0"),
      odata: flags.odata === undefined ? undefined : String(flags.odata),
      cursor,
      scopes: flags.scopes === undefined ? undefined : String(flags.scopes),
      limit: flags.all ? undefined : flags.limit === undefined ? 100 : Number(flags.limit),
      full: !!flags.full,
      profile: selected.profile,
    }, {
      delegated: new DelegatedAuth(new MsalProvider()),
      application: new ApplicationAuth(new MsalApplicationProvider()),
      transport: fetchTransport,
    });
  }
  if (leaf.path === "entra user list" || leaf.path === "entra user show") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    return leaf.path === "entra user list"
      ? listUsers(session, flags, selected.profile, operation, leafHelp(leaf), selected.name)
      : showUser(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
  }
  if (leaf.path === "entra user update") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const readLeaf = LEAVES.find(item => item.path === "entra user show")!;
    const readOperation = operationFor(readLeaf, "v1.0");
    if (!readOperation || readOperation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${readLeaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const coordinator = createMutationCoordinator({
      profile: selected.profile,
      delegated,
      application,
      transport: overrides.mutationTransport ?? mutationFetchTransport,
      ...(overrides.journalPath !== undefined ? { journalPath: overrides.journalPath } : {}),
    });
    return updateUserAccount({
      session,
      coordinator,
      flags,
      profile: selected.profile,
      profileName: selected.name,
      readOperation,
      help: leafHelp(leaf),
    });
  }
  if (leaf.path === "entra user revoke-sessions") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "POST") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const readLeaf = LEAVES.find(item => item.path === "entra user show")!;
    const readOperation = operationFor(readLeaf, "v1.0");
    if (!readOperation || readOperation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${readLeaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const coordinator = createMutationCoordinator({
      profile: selected.profile,
      delegated,
      application,
      transport: overrides.mutationTransport ?? mutationFetchTransport,
      ...(overrides.journalPath !== undefined ? { journalPath: overrides.journalPath } : {}),
    });
    return revokeUserSessions({
      session,
      coordinator,
      flags,
      profile: selected.profile,
      profileName: selected.name,
      readOperation,
      help: leafHelp(leaf),
    });
  }
  if (leaf.path === "entra conditional-access policy update") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const readLeaf = LEAVES.find(item => item.path === "entra conditional-access policy show")!;
    const readOperation = operationFor(readLeaf, "v1.0");
    if (!readOperation || readOperation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${readLeaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const coordinator = createMutationCoordinator({
      profile: selected.profile,
      delegated,
      application,
      transport: overrides.mutationTransport ?? mutationFetchTransport,
      ...(overrides.journalPath !== undefined ? { journalPath: overrides.journalPath } : {}),
    });
    return updateCaPolicy({
      session,
      coordinator,
      flags,
      profile: selected.profile,
      profileName: selected.name,
      readOperation,
      help: leafHelp(leaf),
    });
  }
  if (leaf.path === "entra user authentication-method list" || leaf.path === "entra registration list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    return leaf.path === "entra user authentication-method list"
      ? listAuthenticationMethods(session, flags, selected.profile, operation, help, selected.name)
      : listRegistrationDetails(session, flags, selected.profile, operation, help, selected.name);
  }
  if (leaf.path === "entra group list" || leaf.path === "entra group show" || leaf.path === "entra group member list" || leaf.path === "entra group member-of list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const template = flags.transitive === true
      ? (() => {
        const alternate = TRANSITIVE_OPERATION[leaf.operation!];
        if (!alternate) throw new AxiError("--transitive is available for group member and member-of lists only", "VALIDATION_ERROR", [leafHelp(leaf)]);
        return alternate;
      })()
      : leaf.operation!;
    const operation = operationFor({ ...leaf, operation: template }, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    if (leaf.path === "entra group list") return listGroups(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
    if (leaf.path === "entra group show") return showGroup(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
    return leaf.path === "entra group member list"
      ? listGroupMembers(session, flags, selected.profile, operation, leafHelp(leaf), selected.name)
      : listGroupMemberOf(session, flags, selected.profile, operation, leafHelp(leaf), selected.name);
  }
  if (leaf.path === "entra group member add") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "POST") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    return addGroupMember({
      session,
      profile: selected.profile,
      delegated,
      application,
      transport: overrides.mutationTransport ?? mutationFetchTransport,
      flags,
      profileName: selected.name,
      help: leafHelp(leaf),
    });
  }
  if (leaf.path === "entra directory-role list" || leaf.path === "entra directory-role show" || leaf.path === "entra role-assignment list" || leaf.path === "entra pim eligible list" || leaf.path === "entra pim active list"
    || leaf.path === "entra device list" || leaf.path === "entra device show"
    || leaf.path === "entra administrative-unit list" || leaf.path === "entra administrative-unit show"
    || leaf.path === "entra administrative-unit member list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra directory-role list": return listDirectoryRoles(session, flags, selected.profile, operation, help, selected.name);
      case "entra directory-role show": return showDirectoryRole(session, flags, selected.profile, operation, help, selected.name);
      case "entra role-assignment list": return listRoleAssignments(session, flags, selected.profile, operation, help, selected.name);
      case "entra pim eligible list": return listPimEligible(session, flags, selected.profile, operation, help, selected.name);
      case "entra pim active list": return listPimActive(session, flags, selected.profile, operation, help, selected.name);
      case "entra device list": return listDevices(session, flags, selected.profile, operation, help, selected.name);
      case "entra device show": return showDevice(session, flags, selected.profile, operation, help, selected.name);
      case "entra administrative-unit list": return listAdministrativeUnits(session, flags, selected.profile, operation, help, selected.name);
      case "entra administrative-unit show": return showAdministrativeUnit(session, flags, selected.profile, operation, help, selected.name);
      default: return listAdministrativeUnitMembers(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra organization list" || leaf.path === "entra organization show"
    || leaf.path === "entra organization branding show"
    || leaf.path === "entra organization branding-localization list" || leaf.path === "entra organization branding-localization show") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra organization list": return listOrganizations(session, flags, selected.profile, operation, help, selected.name);
      case "entra organization show": return showOrganization(session, flags, selected.profile, operation, help, selected.name);
      case "entra organization branding show": return showBranding(session, flags, selected.profile, operation, help, selected.name);
      case "entra organization branding-localization list": return listBrandingLocalizations(session, flags, selected.profile, operation, help, selected.name);
      default: return showBrandingLocalization(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra sign-in list" || leaf.path === "entra sign-in show" || leaf.path === "entra directory-audit list" || leaf.path === "entra directory-audit show" || leaf.path === "entra application list" || leaf.path === "entra application show" || leaf.path === "entra service-principal list" || leaf.path === "entra service-principal show" || leaf.path === "entra application owner list" || leaf.path === "entra service-principal owner list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    if ((leaf.path.startsWith("entra sign-in ") || leaf.path.startsWith("entra directory-audit ")) && flags.cursor !== undefined) flags.cursor = (await readCursor(String(flags.cursor)))!;
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra sign-in list": return listSignIns(session, flags, selected.profile, operation, help, selected.name);
      case "entra sign-in show": return showSignIn(session, flags, selected.profile, operation, help, selected.name);
      case "entra directory-audit list": return listDirectoryAudits(session, flags, selected.profile, operation, help, selected.name);
      case "entra directory-audit show": return showDirectoryAudit(session, flags, selected.profile, operation, help, selected.name);
      case "entra application list": return listApplications(session, flags, selected.profile, operation, help, selected.name);
      case "entra application show": return showApplication(session, flags, selected.profile, operation, help, selected.name);
      case "entra service-principal list": return listServicePrincipals(session, flags, selected.profile, operation, help, selected.name);
      case "entra service-principal show": return showServicePrincipal(session, flags, selected.profile, operation, help, selected.name);
      case "entra application owner list": return listApplicationOwners(session, flags, selected.profile, operation, help, selected.name);
      default: return listServicePrincipalOwners(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra service-principal oauth2-grant list" || leaf.path === "entra service-principal app-role-assignment list") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }

    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    return leaf.path === "entra service-principal oauth2-grant list"
      ? listOAuth2Grants(session, flags, selected.profile, operation, help, selected.name)
      : listAppRoleAssignments(session, flags, selected.profile, operation, help, selected.name);
  }
  if (leaf.path === "entra risky-user dismiss") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const readLeaf = LEAVES.find(item => item.path === "entra risky-user show")!;
    const readOperation = operationFor(readLeaf, "v1.0");
    if (!readOperation || readOperation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${readLeaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const coordinator = createMutationCoordinator({
      profile: selected.profile,
      delegated,
      application,
      transport: overrides.mutationTransport ?? mutationFetchTransport,
      ...(overrides.journalPath !== undefined ? { journalPath: overrides.journalPath } : {}),
    });
    return dismissRiskyUser({
      session,
      coordinator,
      flags,
      profile: selected.profile,
      profileName: selected.name,
      readOperation,
      help: leafHelp(leaf),
    });
  }
  if (leaf.path === "entra risky-user list" || leaf.path === "entra risky-user show" || leaf.path === "entra risk-detection list" || leaf.path === "entra risk-detection show") {
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    if (flags.cursor !== undefined) flags.cursor = (await readCursor(String(flags.cursor)))!;
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra risky-user list": return listRiskyUsers(session, flags, selected.profile, operation, help, selected.name);
      case "entra risky-user show": return showRiskyUser(session, flags, selected.profile, operation, help, selected.name);
      case "entra risk-detection list": return listRiskDetections(session, flags, selected.profile, operation, help, selected.name);
      default: return showRiskDetection(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra domain list" || leaf.path === "entra domain show"
    || leaf.path === "entra domain verification-dns-record list" || leaf.path === "entra domain verification-dns-record show"
    || leaf.path === "entra domain service-configuration-record list" || leaf.path === "entra domain service-configuration-record show"
    || leaf.path === "entra domain-dns-record list" || leaf.path === "entra domain-dns-record show") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Domain reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra domain list": return listDomains(session, flags, selected.profile, operation, help, selected.name);
      case "entra domain show": return showDomain(session, flags, selected.profile, operation, help, selected.name);
      case "entra domain verification-dns-record list": return listVerificationDnsRecords(session, flags, selected.profile, operation, help, selected.name);
      case "entra domain verification-dns-record show": return showVerificationDnsRecord(session, flags, selected.profile, operation, help, selected.name);
      case "entra domain service-configuration-record list": return listServiceConfigurationRecords(session, flags, selected.profile, operation, help, selected.name);
      case "entra domain service-configuration-record show": return showServiceConfigurationRecord(session, flags, selected.profile, operation, help, selected.name);
      case "entra domain-dns-record list": return listDomainDnsRecords(session, flags, selected.profile, operation, help, selected.name);
      default: return showDomainDnsRecord(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra certificate-auth-pki list" || leaf.path === "entra certificate-auth-pki show" || leaf.path === "entra certificate-auth-pki count"
    || leaf.path === "entra certificate-authority list" || leaf.path === "entra certificate-authority show" || leaf.path === "entra certificate-authority count") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Certificate-auth reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra certificate-auth-pki list": return listCertificateAuthPkis(session, flags, selected.profile, operation, help, selected.name);
      case "entra certificate-auth-pki show": return showCertificateAuthPki(session, flags, selected.profile, operation, help, selected.name);
      case "entra certificate-auth-pki count": return countCertificateAuthPkis(session, flags, selected.profile, operation, help, selected.name);
      case "entra certificate-authority list": return listCertificateAuthorities(session, flags, selected.profile, operation, help, selected.name);
      case "entra certificate-authority show": return showCertificateAuthority(session, flags, selected.profile, operation, help, selected.name);
      default: return countCertificateAuthorities(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra subscription list" || leaf.path === "entra subscription show" || leaf.path === "entra subscription count") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Subscription reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra subscription list": return listSubscriptions(session, flags, selected.profile, operation, help, selected.name);
      case "entra subscription show": return showSubscription(session, flags, selected.profile, operation, help, selected.name);
      default: return countSubscriptions(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra on-premises-synchronization list" || leaf.path === "entra on-premises-synchronization show") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("On-premises-synchronization reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra on-premises-synchronization list": return listSynchronizations(session, flags, selected.profile, operation, help, selected.name);
      default: return showSynchronization(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra agreement list" || leaf.path === "entra agreement show"
    || leaf.path === "entra agreement acceptance list" || leaf.path === "entra agreement acceptance show"
    || leaf.path === "entra agreement-acceptance list" || leaf.path === "entra agreement-acceptance show") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Terms-of-use reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra agreement list": return listAgreements(session, flags, selected.profile, operation, help, selected.name);
      case "entra agreement show": return showAgreement(session, flags, selected.profile, operation, help, selected.name);
      case "entra agreement acceptance list": return listAgreementAcceptances(session, flags, selected.profile, operation, help, selected.name);
      case "entra agreement acceptance show": return showAgreementAcceptance(session, flags, selected.profile, operation, help, selected.name);
      case "entra agreement-acceptance list": return listAcceptances(session, flags, selected.profile, operation, help, selected.name);
      default: return showAcceptance(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra group-lifecycle-policy list" || leaf.path === "entra group-lifecycle-policy show" || leaf.path === "entra group-lifecycle-policy count"
    || leaf.path === "entra group-setting-template list" || leaf.path === "entra group-setting-template show" || leaf.path === "entra group-setting-template count") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Group lifecycle policy and setting-template reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra group-lifecycle-policy list": return listLifecyclePolicies(session, flags, selected.profile, operation, help, selected.name);
      case "entra group-lifecycle-policy show": return showLifecyclePolicy(session, flags, selected.profile, operation, help, selected.name);
      case "entra group-lifecycle-policy count": return countLifecyclePolicies(session, flags, selected.profile, operation, help, selected.name);
      case "entra group-setting-template list": return listSettingTemplates(session, flags, selected.profile, operation, help, selected.name);
      case "entra group-setting-template show": return showSettingTemplate(session, flags, selected.profile, operation, help, selected.name);
      default: return countSettingTemplates(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra attribute-set list" || leaf.path === "entra attribute-set show" || leaf.path === "entra attribute-set count"
    || leaf.path === "entra custom-security-attribute-definition list" || leaf.path === "entra custom-security-attribute-definition show" || leaf.path === "entra custom-security-attribute-definition count"
    || leaf.path === "entra allowed-value list" || leaf.path === "entra allowed-value show" || leaf.path === "entra allowed-value count") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Custom-security-attribute reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra attribute-set list": return listAttributeSets(session, flags, selected.profile, operation, help, selected.name);
      case "entra attribute-set show": return showAttributeSet(session, flags, selected.profile, operation, help, selected.name);
      case "entra attribute-set count": return countAttributeSets(session, flags, selected.profile, operation, help, selected.name);
      case "entra custom-security-attribute-definition list": return listCustomSecurityAttributeDefinitions(session, flags, selected.profile, operation, help, selected.name);
      case "entra custom-security-attribute-definition show": return showCustomSecurityAttributeDefinition(session, flags, selected.profile, operation, help, selected.name);
      case "entra custom-security-attribute-definition count": return countCustomSecurityAttributeDefinitions(session, flags, selected.profile, operation, help, selected.name);
      case "entra allowed-value list": return listAllowedValues(session, flags, selected.profile, operation, help, selected.name);
      case "entra allowed-value show": return showAllowedValue(session, flags, selected.profile, operation, help, selected.name);
      default: return countAllowedValues(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra contract list" || leaf.path === "entra contract show" || leaf.path === "entra contract count") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Contract reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra contract list": return listContracts(session, flags, selected.profile, operation, help, selected.name);
      case "entra contract show": return showContract(session, flags, selected.profile, operation, help, selected.name);
      default: return countContracts(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  if (leaf.path === "entra identity-provider list" || leaf.path === "entra identity-provider show" || leaf.path === "entra identity-provider count" || leaf.path === "entra identity-provider available-types") {
    if (String(flags["api-version"] ?? "v1.0") !== "v1.0") {
      throw new AxiError("Identity-provider reads support v1.0 only; beta needs its own review", "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    const selected = profiles.resolve(flags.profile as string | undefined);
    const operation = operationFor(leaf, "v1.0");
    if (!operation || operation.method !== "GET") {
      throw new AxiError(`Unknown catalogued Graph operation for ${leaf.path}`, "VALIDATION_ERROR", [leafHelp(leaf)]);
    }
    let delegated = overrides.delegated;
    let application = overrides.application;
    if (!delegated) {
      const { DelegatedAuth: Service } = await import("./auth.js");
      const { MsalProvider } = await import("./msal-provider.js");
      delegated = new Service(new MsalProvider());
    }
    if (!application) {
      const { ApplicationAuth: Service } = await import("./app-auth.js");
      const { MsalApplicationProvider } = await import("./msal-app-provider.js");
      application = new Service(new MsalApplicationProvider());
    }
    const session = new GraphSession({ delegated, application, transport: overrides.transport ?? fetchTransport });
    const help = leafHelp(leaf);
    switch (leaf.path) {
      case "entra identity-provider list": return listIdentityProviders(session, flags, selected.profile, operation, help, selected.name);
      case "entra identity-provider show": return showIdentityProvider(session, flags, selected.profile, operation, help, selected.name);
      case "entra identity-provider available-types": return availableIdentityProviderTypes(session, flags, selected.profile, operation, help);
      default: return countIdentityProviders(session, flags, selected.profile, operation, help, selected.name);
    }
  }
  const operation = operationFor(leaf, String(flags["api-version"] ?? "v1.0"));
  throw new AxiError(`Command is not executable: ${operation?.disposition ?? "unavailable"} (${operation?.owningSlice ?? "no inventory mapping"})`, "NOT_IMPLEMENTED", [leafHelp(leaf)]);
}

export async function main() {
  await runAxiCli({
    description: DESCRIPTION,
    version: VERSION,
    topLevelHelp: TOP_LEVEL_HELP,
    // Strict resolution owns leaf help, so SDK help cannot bypass validation.
    argv: ["dispatch"],
    home: () => localHome(),
    commands: {
      dispatch: async () => {
        if (process.argv.length === 3 && process.argv[2] === "--help") return TOP_LEVEL_HELP;
        return executeArgv(process.argv.slice(2));
      },
    },
    formatError: error => ({
      output: `${encode(error instanceof AxiError ? { error: error.message, code: error.code, help: error.suggestions } : { error: "Unable to run mg-axi", help: ["mg-axi --help"] })}\n`,
      exitCode: error instanceof AxiError && error.code === "VALIDATION_ERROR" ? 2 : 1,
    }),
  });
}
