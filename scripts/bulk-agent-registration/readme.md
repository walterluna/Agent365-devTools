# Agent 365 bulk onboarding script implementation

For instructions to configure and run bulk Agent 365 onboarding, see Microsoft Learn: [Build with answers in reach](https://learn.microsoft.com/).

This README describes the implementation and architecture of the PowerShell scripts in this directory. It's intended for contributors who need to understand, maintain, troubleshoot, or extend the scripts. Customer prerequisites, permission setup, CSV authoring, validation commands, run examples, and result-checking instructions belong in Microsoft Learn, not in this implementation reference.

| Item | Value |
| --- | --- |
| Document scope | Script responsibilities, dependency flow, authentication, Graph operations, and reporting internals |
| Scripts covered | 15 - the orchestrator; the 4 provisioning and 4 removal scripts it invokes; `New-A365AutomationApp.ps1` as a one-time prerequisite; the 4 dedicated `Update-A365*.ps1` entry points; and `A365-BulkOnboarding.ps1` for CSV-driven bulk runs (`A365-BulkOnboardingCsv.psm1` is its supporting module, not a standalone script) |
| API surface | Microsoft Graph only, with one documented exception (Azure Key Vault) |
| Minimum PowerShell | `7.0 (#requires -Version 7)` |

## 1. Architecture and orchestration

`A365-AutomationOrchestrator.ps1` provisions a complete Microsoft Agent 365 agent by invoking a set of single-purpose scripts in dependency order. It performs no Graph calls of its own — it validates the request, decides which phases run, and passes the identifiers each phase produces to the next one.

### 1.1 The four objects

| # | Object | What it is | Created by |
| --- | --- | --- | --- |
| 1 | Blueprint | An Entra application of type agentIdentityBlueprint, plus its blueprint principal. Shared by every agent built from it. | `New-A365AgentBlueprint.ps1` |
| 2 | Agent identity | A servicePrincipal of type agentIdentity, built on the blueprint. This is the agent's identity for token acquisition (fmi_path). | `New-A365AgentIdentity.ps1` |
| 3 | Agent user | A real directory user of type agentUser, bound to the agent identity. Can hold a license and a mailbox. | `New-A365AgentUser.ps1` |
| 4 | Registration | The Microsoft 365 admin center inventory entry plus the agent registry instance. | `New-A365AgentRegistration.ps1` |

> **Note: Agent user and registration are siblings, not a chain.** Both depend only on the agent identity. If the agent user phase fails, the registration phase still runs. That is deliberate — a licensing problem should not prevent the agent from being registered.

### 1.2 Three modes of operation

| Mode | Switches | Behavior |
| --- | --- | --- |
| Create | `-NewBlueprint, -NewAgentIdentity, -NewAgentUser, -NewAgentRegistration` | Creates the selected objects. Phases run in dependency order 1→2→3→4. |
| Update | `-UpdateBlueprint, -UpdateAgentIdentity, -UpdateAgentUser, -UpdateAgentRegistration` | Each takes the id of the object to change and writes ONLY the attributes you supply. |
| Remove | `-RemoveBlueprint, -RemoveAgentIdentity, -RemoveAgentUser, -RemoveAgentRegistration` | Each takes the id of the object to delete. Runs AFTER all create/update phases, in reverse dependency order. |

Create, update and remove can be combined in a single run. Removal always runs last, and in reverse dependency order (registration → agent user → agent identity → blueprint) so a parent is never deleted before the children discovered through it.

Phase-prefixed parameters are routed to the corresponding child script. Parameters for unselected phases are reported as ignored. Existing blueprint and identity identifiers can seed later phases without creating their parents.

### 1.3 Full capability list

- Create a blueprint with sponsors, owners, description, and declared API permissions.
- Grant tenant admin consent on the blueprint principal and make permissions inheritable.
- Create a client secret on the blueprint, and optionally store it in Azure Key Vault.
- Federate a managed identity to the blueprint (preferred over a client secret in production).
- Create an agent identity on a new or existing blueprint, with sponsors, owners and tags.
- Assign custom security attributes to the agent identity, with pre-flight validation.
- Create an agent user with manager, usage location, and an assigned license.
- Register the agent in the Microsoft 365 admin center and the agent registry.
- Update any of the four objects, writing only the attributes supplied.
- Remove any of the four objects, with an inspect-only dry run and an optional recycle-bin purge.
- Write a per-run log for the orchestrator and every script it invokes, with credential redaction on by default.
- Emit a machine-readable JSON report of everything created, changed, or left outstanding.

## 2. Runtime dependencies and script resolution

The scripts declare `#requires -Version 7`. Graph operations use `Invoke-MgGraphRequest` and an authenticated session supplied by `Microsoft.Graph.Authentication`, rather than the full Graph SDK module set. No a365 CLI is invoked.

Interactive Key Vault writes additionally obtain a vault-audience token from an Azure CLI or Az.Accounts session. App-only Key Vault token acquisition does not depend on either tool (section 6).

The orchestrator resolves step scripts relative to its own directory unless `-ScriptRoot` overrides it. Required scripts are checked before execution; a missing script causes an immediate failure identifying the file and phase.

## 3. Graph operations and permission mappings

Unattended runs authenticate as an Entra application. `New-A365AutomationApp.ps1` creates that application, adds the Microsoft Graph application roles for the scenarios you select, and can grant admin consent in the same run.

### 3.1 Automation application responsibilities

`New-A365AutomationApp.ps1` grants app roles by default; `-SkipGrant` suppresses the grant without suppressing application or credential creation. If consent fails, it reports the outstanding grants and consent link. Scenario role sets are additive:

| Scenario | Covers |
| --- | --- |
| `Blueprint` | Creating and configuring blueprints and blueprint principals |
| `AgentIdentity` | Creating agent identities, including tags and custom security attributes |
| `Registration` | Registering agents in the admin center and the agent registry |
| `AgentUser` | Agent users, licensing, and per-identity permission consent |
| `All` | Every role in the four scenarios above |

### 3.2 Microsoft Graph application permissions

The mappings below explain which operations in each provisioning script require Microsoft Graph application roles. They are an implementation reference, not a customer permission-configuration checklist. Application roles authorize calls only after consent on the automation service principal.

#### Blueprint: `New-A365AgentBlueprint.ps1`

| Permission | Why it is needed |
| --- | --- |
| `AgentIdentityBlueprint.Create` | Create the blueprint application |
| `AgentIdentityBlueprint.Read.All` | Read blueprints, resolve one by appId or objectId |
| `AgentIdentityBlueprint.ReadWrite.All` | Update the blueprint, declare required permissions, set inheritance |
| `AgentIdentityBlueprint.AddRemoveCreds.All` | addPassword (client secret) and federated identity credentials |
| `AgentIdentityBlueprintPrincipal.Create` | Create the blueprint principal |
| `AgentIdentityBlueprintPrincipal.Read.All` | Read the blueprint principal |
| `AgentIdentityBlueprintPrincipal.ReadWrite.All` | Assign owners on the blueprint PRINCIPAL |
| `Application.Read.All` | Resolve resource service principals when declaring permissions |
| `Application.ReadWrite.All` | Assign owners on the blueprint APPLICATION |
| `Directory.Read.All` | Resolve directory objects for owner assignment |
| `User.Read.All` | Resolve sponsors and owners given as UPN, mail or display name |
| `Group.Read.All` | Resolve group sponsors |

#### Agent identity: `New-A365AgentIdentity.ps1`

| Permission | Why it is needed |
| --- | --- |
| `AgentIdentity.Create.All` | Create the agent identity |
| `AgentIdentity.Read.All` | Read agent identities and verify the created object |
| `AgentIdentity.ReadWrite.All` | Update tags, display name, sponsors and owners |
| `AgentIdentityBlueprint.Read.All` | Validate the blueprint the identity is built on |
| `Application.Read.All` | Resolve resource service principals |
| `Application.ReadWrite.All` | Assign owners on the agent identity |
| `Directory.Read.All` | Resolve directory objects |
| `User.Read.All` | Resolve sponsors and owners |
| `Group.Read.All` | Resolve group sponsors |
| `CustomSecAttributeAssignment.ReadWrite.All` | Assign custom security attributes |
| `CustomSecAttributeDefinition.Read.All` | Validate attribute sets and allowed values before assigning |

*The last two rows are requested twice over: as application roles (shown here) and, for `-Scenario AgentIdentity` and `-Scenario All`, as delegated scopes as well. An unattended app-only run needs only the application roles. An interactive run needs the delegated scopes AND the Attribute Assignment Administrator directory role, which no application permission grants.*

#### Registration: `New-A365AgentRegistration.ps1`

| Permission | Why it is needed |
| --- | --- |
| `AgentRegistration.ReadWrite.All` | POST /beta/copilot/agentRegistrations |
| `AgentInstance.ReadWrite.All` | POST /beta/agentRegistry/agentInstances |
| `AgentIdentity.Read.All` | Validate the agent identity being registered |
| `User.Read.All` | Resolve `-AgentRegistrationOwner` to object ids (app-only cannot use /me) |

#### Agent user: `New-A365AgentUser.ps1`

*This set covers the agent user phase of the orchestrator. It holds the only copy of `AgentIdUser.ReadWrite.All`, which is the role that authorizes creating an agent user, so it is also required by `-Scenario All`.*

| Permission | Why it is needed |
| --- | --- |
| `AgentIdUser.ReadWrite.All` | Create the agent user. This is the specific role the call is authorized by. |
| `User.ReadWrite.All` | Set manager and usage location, assign the license |
| `User.Read.All` | Resolve the manager and check for an existing account |
| `Directory.Read.All` | Read verified domains, resolve the manager |
| `Organization.Read.All` | Read /subscribedSkus for the license pre-flight check |
| `DelegatedPermissionGrant.ReadWrite.All` | Per-identity delegated consent |
| `AppRoleAssignment.ReadWrite.All` | Per-identity application role assignment |
| `Application.Read.All` | Resolve resource service principals |
| `AgentIdentity.Read.All` | Validate the identity the user is bound to |

### 3.3 Three permissions that are commonly missed

Each of these fails with a bare 403 that names no permission, so the cause is not obvious from the error:

| Permission | The trap |
| --- | --- |
| `AgentIdentityBlueprintPrincipal.ReadWrite.All` | Required to set owners on the blueprint PRINCIPAL. `Application.ReadWrite.All` does NOT authorize it — granting more `Application.*` permissions will not fix a principal owner denial. |
| `AgentIdUser.ReadWrite.All` | Required to create an agent user. `User.ReadWrite.All` is NOT enough. |
| `CustomSecAttributeAssignment.ReadWrite.All` | Custom security attributes are gated separately from every other directory permission. No `Application.*` or `Directory.*` role reaches them, and Global Administrator does not hold them by default. `New-A365AutomationApp.ps1 -Scenario AgentIdentity` (or `All`) requests them in BOTH shapes: as application roles for an unattended run, and as delegated scopes for an interactive one. An interactive run needs one more thing that no application permission can grant - the signed-in user must also hold the Attribute Assignment Administrator directory role. |

## 4. Consent failure handling

Consent and app-role assignment have a separate authorization boundary from object creation. The scripts' failure messages distinguish the directory roles that can grant consent:

| What you are consenting | Directory role required |
| --- | --- |
| Microsoft Graph APPLICATION roles (application permissions) — this is most of the table in section 3 | Privileged Role Administrator, or Global Administrator |
| Anything else: delegated scopes, or application roles on a non-Graph API | Application Administrator, Cloud Application Administrator, Privileged Role Administrator, or Global Administrator |

> **Important: Application Administrator cannot consent Microsoft Graph app roles.** Application Administrator and Cloud Application Administrator may consent any permission for any API EXCEPT Microsoft Graph application permissions. Because the A365 roles are Graph application permissions, consenting them needs Privileged Role Administrator or Global Administrator.

A refused consent or app-role assignment is recorded rather than undoing an object that has already been created. Child results carry `adminConsentUrl`, `portalPermissionsUrl` and `consentFailures`; the orchestrator aggregates outstanding work under `summary.consentActionRequired`.

## 5. Authentication implementation

Authentication validation distinguishes credential-based app-only connections from delegated interactive connections and rejects conflicting methods before phase execution.

| Method | Parameters | App-only? | Notes |
| --- | --- | --- | --- |
| Client secret | `-ClientId -ClientSecret` | Yes | Omitted secrets fall back to `A365_CLIENT_SECRET`. |
| Certificate | `-ClientId -CertificateThumbprint`<br>`-Certificate \| -CertificatePath` | Yes | Supports a store lookup, a live certificate object, or a file with `-CertificatePassword`. |
| Managed identity | `-UseManagedIdentity` | Yes | `-ClientId` selects a user-assigned identity. |
| Access token | `-AccessToken` | Depends | A pre-obtained Microsoft Graph token. |
| Interactive | `-Interactive` | No | Signs in as a user. Cannot be used with `-NewAgentUser`. |

> **Important: `-NewAgentUser` requires app-only authentication.** `New-A365AgentUser.ps1` supports client secret, certificate and managed identity only — it has no interactive mode. The orchestrator refuses the combination before any phase runs, rather than failing after the blueprint and identity have already been created.

### 5.1 Mixed authentication for the registration phase

`-AgentRegistrationAuth Same` reuses the run's authentication. `Interactive` replaces the registration phase's authentication with a delegated connection while earlier phases retain their app-only credentials. This separates registration authorization from the app-only path required to create an agent user.

The interactive registration argument map retains `ClientId` and `SkipPermissionCheck` when applicable, but not the app-only credentials. The AgentUser argument map filters out `Interactive` and `SkipPermissionCheck`, which its script does not declare. Removal scripts receive only their supported authentication subset: `ClientId`, `ClientSecret`, `CertificateThumbprint`, `AccessToken`, and `Interactive`.

### 5.2 Credential forwarding and redaction

Authentication parameters are passed to child scripts as PowerShell objects, not serialized command-line text. Client secrets, certificate passwords and access tokens accept strings or `SecureString`; certificates can remain live `X509Certificate2` objects. The update wrappers preserve the same objects when forwarding to their provisioning scripts.

If `-ClientSecret` is omitted, the scripts read `A365_CLIENT_SECRET`. Plain-text credentials supplied as parameters trigger an exposure warning. Log and report redaction are separate paths, described in sections 7 and 10.

## 6. Azure Key Vault implementation

`-BlueprintKeyVaultName` stores a newly created blueprint client secret in Azure Key Vault instead of printing it once to the console. This is the only capability in the suite that does not run on Microsoft Graph.

### 6.1 Why this one is not a Graph operation

Key Vault is a separate service with its own data plane and its own token audience. This was verified against a live tenant rather than assumed:

| Probe | Result | Conclusion |
| --- | --- | --- |
| `GET /beta/keyVaults, /v1.0/keyVaults, /beta/secrets` | `400` | Microsoft Graph exposes no Key Vault surface at all |
| Graph token → vault data plane | `401` | The vault rejects a Graph token outright |
| Vault-audience token → the same request | `403` | Authenticated; only the RBAC role was missing |
| The same Graph token → /v1.0/organization | `200` | Control: the Graph token was valid |

The 401/403 split is what makes this conclusive. Because the vault answered an authorization question when given a vault-audience token, the 401 was purely about audience. A token is bound to its audience and cannot be exchanged.

### 6.2 Azure RBAC authorization boundary

> **Important: Owner and Contributor cannot read or write secrets.** Key Vault separates its management plane (create and configure vaults) from its data plane (read and write secrets). Owner and Contributor grant the former and have NO dataActions at all. This was confirmed live: a caller holding Contributor on the vault still received 403 on a secret write.

| Role | Role definition id | Grants |
| --- | --- | --- |
| Key Vault Secrets Officer | `b86a8fe4-44ce-4948-aee5-eccb2c155cd7` | `Microsoft.KeyVault/vaults/secrets/*` — required to WRITE |
| Key Vault Secrets User | `4633458b-17de-408a-b874-0445c86b69e6` | getSecret + readMetadata — read secret values |
| Key Vault Reader | `21090545-7ca7-4776-b22c-e363652d74d2` | List keys and secrets, but NOT read secret values |
| Key Vault Crypto Officer | `14b46e9e-c2b7-41b4-b07b-48a6ebf60603` | `Microsoft.KeyVault/vaults/keys/*` — manage keys |
| Key Vault Administrator | `00482a5a-887f-4fb3-b363-3b7fe8e74483` | All data actions: keys, secrets and certificates |

Vault authorization applies to the automation application's service principal object id, not its application object id. `New-A365AutomationApp.ps1` grants Graph roles and cannot grant Azure RBAC access. Role-assignment writes have their own authorization boundary: Owner, User Access Administrator, or Role Based Access Control Administrator can perform them; Contributor excludes `Microsoft.Authorization/*/Write`.

For vaults using access policies instead of RBAC, the corresponding data-plane operations require secret set and get permissions. The script reports which authorization model it encountered.

### 6.3 How the vault token is obtained

| Graph authentication | Key Vault token |
| --- | --- |
| `-ClientSecret` | Client credentials against the token endpoint, with the vault scope |
| `-Certificate / -CertificateThumbprint` | A signed RS256 client assertion — the Graph SDK will not mint a token for another audience |
| `-UseManagedIdentity` | IMDS, or IDENTITY_ENDPOINT on App Service and Functions |
| `-Interactive` | A signed-in Azure session (az login or Connect-AzAccount) |
| `-AccessToken` | A Graph token cannot authorize the vault request; a separate `-KeyVaultAccessToken` is needed |

### 6.4 What gets stored

The secret is written as a new VERSION (an existing name is never overwritten, so credential history is preserved), with the vault entry expiring at the same moment as the Entra credential itself:

| Field | Value |
| --- | --- |
| Secret name | The display name folded to the characters Key Vault allows, or `-BlueprintKeyVaultSecretName` |
| contentType | `application/x-a365-client-secret` |
| Tags | `a365Object, a365AppId, a365DisplayName` |
| Expiry | Aligned with the credential lifetime, so a stale secret is not left looking current |

The write is confirmed by reading the version back. If the vault write fails for any reason, the secret is still printed to the console — a credential that exists in Entra but was captured nowhere is unrecoverable.

## 7. Logging and JSON implementation

> **Security: Bearer tokens are always redacted.** Credential redaction is on by default. `-LogIncludeSecrets` opts in to recording plain-string client secrets and passwords, but SecureString values, tokens, and anything JWT-shaped are removed unconditionally — a captured token is directly replayable until it expires.

### 7.1 Log files

`-LogPath` produces one log per script per run. Every file in a run shares a correlation id, and the orchestrator log ends with an index naming each child log and the phase that produced it — including grandchildren, such as the cascade a blueprint removal delegates to.

```text
<script-name>-yyyy-MM-dd_HH-mm-ss-<correlationId>.log
```

Windows file names cannot contain a colon, so the time is written HH-mm-ss. Logs are flushed line by line rather than buffered, and a trap writes the run footer even when the script fails, so the log is complete up to the moment of a crash.

Each log records:

- Every Graph call: method, URI, request body, response, HTTP status and duration.
- Retry decisions and tolerated status codes.
- The full argument list handed to each child script.
- Phase banners and per-step results, so the log explains what the run decided as well as what it called.

### 7.2 Orchestrator report

The orchestrator assembles an `A365ProvisioningRunReport`; `-OutputJsonPath` serializes it with schema metadata, `run`, `phases`, `summary`, and `steps`:

- `run` carries invocation context, correlation, and timing.
- `phases` records each script's status, timing, parameters, identifiers, and child result.
- `summary.identifiers` combines blueprint application and principal ids, agent identity id and `fmiPath`, agent user id and UPN, and registration id.
- `summary.consentActionRequired` aggregates outstanding grants with per-object consent links and required roles.
- `summary.tokenRequest` is a token-request template containing the resolved blueprint and identity ids, not a minted token.
- `steps` contains the per-step result table. Blueprint phase results also contribute Key Vault secret metadata and read-back verification.

File serialization redacts blueprint secrets unless explicitly enabled. The orchestrator's in-memory return value retains those secrets, so the bulk wrapper must apply its own redaction before serializing child results.

## 8. Implementation troubleshooting

These are failure modes confirmed against a live tenant. Each produces an error that does not obviously point at its cause.

| Symptom | Implementation cause |
| --- | --- |
| 403 with no permission named, when assigning owners on a blueprint principal | Needs `AgentIdentityBlueprintPrincipal.ReadWrite.All`. `Application.ReadWrite.All` does not authorize that call, so granting more `Application.*` permissions will not help. |
| 403 'Insufficient privileges' when creating an agent user | Needs `AgentIdUser.ReadWrite.All`. `User.ReadWrite.All` is not enough. |
| 403 when assigning custom security attributes | Needs `CustomSecAttributeAssignment.ReadWrite.All` and `CustomSecAttributeDefinition.Read.All`. These are not implied by any `Directory.*` or `Application.*` role, nor held by Global Administrator by default. |
| Consent appears to succeed but the agent gets no claims | Declaring a permission does not grant it; unsuccessful grants are recorded in `consentFailures`. |
| Registration 'Deleted' but the agent is still listed | The registration id has two spellings: a bare GUID and a tenant-prefixed `T_<guid>`. DELETE answers 404 for an unknown id, and 404 is the desired end state, so the wrong spelling reports success. The scripts probe both and report a 404 as 'Already absent' rather than 'Deleted'. |
| Key Vault write returns 401 | A Microsoft Graph token was presented. Key Vault requires a separate audience; token acquisition is described in section 6.3. |
| Key Vault write returns 403 while the caller holds Owner or Contributor | Those are management-plane roles with no dataActions; secret writes are authorized separately. |
| `az role assignment create` fails with `AuthorizationFailed` | Creating role assignments needs Owner, User Access Administrator, or RBAC Administrator. Contributor carries notActions `Microsoft.Authorization/*/Write`. |
| A role assignment was just made but still returns 403 | Azure RBAC enforcement is not immediate; propagation delay and cached authorization context can affect probes. |
| 'Ignored: -X applies to ... which was not selected' | A parameter was supplied for a phase that did not run. Harmless, but usually means a switch was forgotten. |

## 9. Script inventory

The orchestrator invokes the four `New-A365Agent*` provisioning scripts and the four `Remove-A365*` scripts directly, passing identifiers between dependent phases. `New-A365AutomationApp.ps1` is a separate bootstrap entry point; neither the orchestrator nor the bulk wrapper invokes it.

### 9.1 Provisioning

| Script | Purpose |
| --- | --- |
| `A365-AutomationOrchestrator.ps1` | Runs the whole pipeline; makes no Graph calls of its own |
| `New-A365AutomationApp.ps1` | PREREQUISITE, not invoked by the orchestrator. Creates the automation app and grants its Graph permissions. |
| `New-A365AgentBlueprint.ps1` | Creates and configures a blueprint and its principal |
| `New-A365AgentIdentity.ps1` | Creates an agent identity on an existing blueprint |
| `New-A365AgentUser.ps1` | Creates an agent user, with manager and license |
| `New-A365AgentRegistration.ps1` | Registers the agent in the admin center and registry |
| `A365-BulkOnboarding.ps1` | Provisions many agents from one CSV file; drives the orchestrator once per row, in dependency order |

### 9.2 Removal

| Script | Purpose |
| --- | --- |
| `Remove-A365Blueprint.ps1` | Deletes a blueprint and its principal; can cascade with -Force |
| `Remove-A365AgentIdentity.ps1` | Deletes agent identities, and nothing else |
| `Remove-A365AgentUser.ps1` | Deletes agent users, and nothing else |
| `Remove-A365AgentRegistration.ps1` | Deletes the admin center registration, and nothing else |

The removal order is registration, agent user, agent identity, then blueprint. Agent users are discovered through their owning identity, so deleting the identity first can orphan a user that still holds a license and UPN. Blueprint removal can delegate cascading cleanup to the other removal scripts; inspect-only and permanent-purge paths remain separate from creation.

### 9.3 Update wrappers

These four scripts are an alternative interface to the orchestrator's `-UpdateX` parameters, not extra steps. Each forwards to the matching `New-A365*.ps1` in update mode, so both routes execute the same implementation.

| Script | Purpose |
| --- | --- |
| `Update-A365Blueprint.ps1` | Changes an existing blueprint. Takes `-BlueprintId` (application or object id). |
| `Update-A365AgentIdentity.ps1` | Changes an existing agent identity. Takes `-AgentIdentityId`. |
| `Update-A365AgentUser.ps1` | Changes an existing agent user. Takes `-AgentUserId`. |
| `Update-A365AgentRegistration.ps1` | Changes an existing registration. Takes `-RegistrationId`. |

Forwarding is based on bound parameters, not variable defaults: only explicitly supplied attributes are written, so an omitted parameter cannot overwrite stored state.

Merge behavior differs by attribute. Custom security attributes merge per attribute, though a multi-valued attribute that is written is replaced wholesale. Agent identity tags merge by reading the current set and writing the union. Agent identity owners are additive; registration ownerIds replace the whole collection.

Like the step scripts they wrap (section 5.2), `-ClientSecret`, `-CertificatePassword` and `-AccessToken` accept a string or `SecureString` and are forwarded unchanged. Plain strings remain supported but trigger the step script's command-line-exposure warning.

### 9.4 Supporting module and tests

| File | Responsibility |
| --- | --- |
| `A365-BulkOnboardingCsv.psm1` | Pure CSV parsing, validation, dependency ordering, row-to-parameter mapping, execution-state construction, and report redaction; no Graph or orchestrator calls |
| `tests/Run-Tests.ps1` | Runs the local test suite for parsing, validation, ordering, mapping, execution, reporting, secure authentication inputs, and the sample CSV |

## 10. Bulk-processing implementation

For CSV format, validation commands, execution examples, and customer-facing output guidance, see [Microsoft Learn](https://learn.microsoft.com/). This section explains the implementation behind that workflow.

### 10.1 Invocation flow

`A365-AutomationOrchestrator.ps1` provisions one agent per invocation. `A365-BulkOnboarding.ps1` reads a CSV describing many blueprints, agent identities, agent users and registrations, builds a dependency plan, and calls the orchestrator once per row that needs creating. Each call selects one creation phase and receives its parent's resolved id.

The orchestrator is invoked in-process with the PowerShell call operator, in the same runspace as the wrapper script. No new process is started and nothing is serialized: a SecureString client secret or an X509Certificate2 passed to `-ClientSecret` / `-Certificate` stays the exact same live object for every row that uses it.

### 10.2 Parser and parameter mapping

`A365-BulkOnboardingCsv.psm1` reads rows with `Import-Csv -LiteralPath` and preserves an array even for a single-row file. It owns the header allowlist and per-object schema, converting cell values into typed parameters: semicolon-delimited arrays, strict booleans, validated GUIDs and UPNs, enumerations, and JSON parsed as hashtables. Nonblank fields that do not apply to a row's object type are rejected instead of silently ignored.

The schema selects the creation switch, phase-prefixed parameters, and parent-reference parameter. Blueprint ids are forwarded as `UseExistingBlueprint`; identity ids are forwarded as `UseExistingAgentIdentity`. Registration display names cannot inherit a preceding identity row's display name because each row is a separate orchestrator invocation.

`ParameterJson` is an allowlisted extension point for advanced phase parameters, not an arbitrary splat. Full parameter names are required. Tenant, credentials, authentication, parent identifiers, action switches, output, and logging settings are blocked so a row cannot override the execution context or dependency plan.

### 10.3 Validation and dependency ordering

The entire file is validated before execution: required and recognized headers, object types, case-insensitive unique keys, legal `ExistingId` and `ParentKey` combinations, typed values, required creation fields, parent types, and cycles. Errors are collected with row and column context. AgentUser and AgentRegistration are leaf nodes, both parented to AgentIdentity.

An `ExistingId` row is a reference-only Blueprint or AgentIdentity anchor. Its supplied id initializes resolved state; the wrapper neither creates nor updates it and does not invoke the orchestrator for it. Existing anchors are separated from the creation graph, so their children do not wait for a creation phase.

Creation rows are topologically sorted with Kahn's algorithm. Among currently ready rows, CSV row order is the tie-breaker. Each successful parent publishes its identifier into execution state before a dependent row's parameters are assembled.

### 10.4 Execution and failure propagation

CSV parsing and validation precede `ShouldProcess`. A `-WhatIf` run or declined confirmation produces a plan without invoking the orchestrator. For an actual run, authentication is validated once before processing creation rows. AgentUser rows require client-secret, certificate, or managed-identity authentication; interactive and pre-obtained access-token modes are rejected for those rows.

Tenant and authentication inputs are taken from the wrapper's bound parameters and forwarded unchanged to every row invocation. When `LogPath` is bound, logging arguments and the shared correlation id are forwarded separately; the wrapper generates an id when omitted. Registration rows can select mixed authentication through the mapping to `AgentRegistrationAuth` (section 5.1).

A row with an unresolved parent is marked `SkippedDependency` and never invokes the orchestrator. A failed creation is marked `Failed`; descendants are skipped through the same parent-state check. Independent trees continue because failure is recorded per row rather than terminating the processing loop. AgentUser failure does not block its sibling registration. A failed or dependency-skipped run exits with code 1 only after aggregate reporting.

### 10.5 Aggregate report and redaction

The wrapper assembles one `A365BulkProvisioningRunReport`, not a set of per-row files. The orchestrator is never given `-OutputJsonPath`; its return value is captured in memory as the row's child result.

The aggregate includes schema metadata, `run` context and correlation id, `mode`, `totals`, `rows`, and `validationErrors`. Row entries carry dependency keys, status, resolved identifiers, errors, and child results. Existing anchors, created rows, failures, and dependency skips therefore remain visible in the same report; preview mode records creation rows as `Planned`.

Child results are recursively redacted before serialization. Blueprint client secrets are excluded unless `-IncludeBlueprintSecretsInOutput` is enabled; authentication inputs remain redacted regardless of that switch. Before writing a secret-bearing report, the writer attempts to restrict file permissions and warns if it cannot. Logging's `-LogIncludeSecrets` switch is independent of JSON secret inclusion.
