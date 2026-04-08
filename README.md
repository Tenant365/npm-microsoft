# @tenant365/microsoft

Microsoft APIs – Powered by Tenant365

A TypeScript library for Microsoft 365 and Azure authentication. Supports client credentials, certificate-based authentication, and Azure Key Vault-backed signing.

Detailed architecture and module conventions: see `ARCHITECTURE.md`.

---

## Features

- **Client Credentials** – Authenticate with a client secret
- **Certificate Authentication** – Sign JWT assertions locally with a private key
- **Key Vault Signing** – Sign JWT assertions remotely via Azure Key Vault (private key never leaves the vault)
- **Key Vault Utilities** – Fetch certificates and secrets from Azure Key Vault
- **Microsoft Teams Client** – Request all Teams or a specific Team via Microsoft Graph
- **Microsoft Entra (Graph directory)** – Users, groups, applications, service principals, directory roles
- **Dual module output** – CommonJS and ESM, with TypeScript declarations

---

## Installation

```bash
npm install @tenant365/microsoft
# or
pnpm add @tenant365/microsoft
```

---

## Module Structure

The package is organized by domain, each with an object-oriented client and typed helpers:

- `src/core/auth` - authentication and signing strategies
- `src/common/graph` - shared Graph transport + base client abstractions
- `src/teams` - Teams client (`types`, `guards`, `client`, `index`)
- `src/sharepoint` - SharePoint client (`types`, `guards`, `client`, `index`)
- `src/entra` - Entra directory clients (users, groups, applications, service principals, directory roles)

Design principles:

- explicit function names per signing strategy
- runtime guards for Graph payload shape validation
- client classes per domain with factory functions
- shared Graph transport (`src/common/graph/request.ts`) for consistent HTTP/error handling
- Vitest unit tests per module

---

## Quick Start

### Client Credentials (Secret)

```typescript
import { createM365ClientCredentials, MS365Scopes } from "@tenant365/microsoft";

const client = createM365ClientCredentials({
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  clientSecret: "your-client-secret",
});

const { token, expiresAt } = await client.GetAccessToken(MS365Scopes.DEFAULT);
// Use `token` as a Bearer token for Microsoft Graph or other APIs
```

### Certificate Authentication (Local Private Key)

```typescript
import { createM365ClientCertificate, MS365Scopes } from "@tenant365/microsoft";

const client = createM365ClientCertificate({
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  privateKey: "-----BEGIN PRIVATE KEY-----\n...",
  certificate: "-----BEGIN CERTIFICATE-----\n...",
});

const { token, expiresAt } = await client.GetAccessToken(MS365Scopes.DEFAULT);
```

You can generate local PEM material with:

```typescript
import { createM365LocalSigningCertificate } from "@tenant365/microsoft";

const cert = await createM365LocalSigningCertificate({
  commonName: "my-local-app",
  daysValid: 365,
});
// cert.privateKeyPem
// cert.certificatePem
```

### Key Vault Signing

The private key never leaves Azure Key Vault. The library fetches the certificate and delegates signing to the Key Vault sign API.

```typescript
import {
  getM365AccessTokenWithKeyVaultSigning,
  MS365Scopes,
} from "@tenant365/microsoft";

const { token, expiresAt } = await getM365AccessTokenWithKeyVaultSigning({
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  clientSecret: "your-keyvault-client-secret",
  keyVaultName: "your-vault-name",
  certificateName: "your-certificate-name",
  keyName: "your-key-name",
  keyVaultTenantId: "your-keyvault-tenant-id", // optional, defaults to tenantId
  keyVaultClientId: "your-keyvault-client-id", // optional, defaults to clientId
  keyVaultClientSecret: "your-keyvault-client-secret", // optional, defaults to clientSecret
  scope: MS365Scopes.DEFAULT,
});
```

### Teams Client

```typescript
import {
  createM365ClientCredentials,
  MS365Scopes,
  createTeamsClient,
} from "@tenant365/microsoft";

const auth = createM365ClientCredentials({
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  clientSecret: "your-client-secret",
});

const teamsClient = createTeamsClient(auth);

const teams = await teamsClient.getAllTeams();
const teamsSearch = await teamsClient.getTeamsBySearch("keyword");
const metadata = await teamsClient.getAllTeamsMetadata();
const team = await teamsClient.getTeam("team-id");
const sameTeam = await teamsClient.getTeamById("team-id");
const token = await teamsClient.getAccessToken(); // uses MS365Scopes.DEFAULT

// Create team: Graph requires template + at least one owner (user Azure AD object id)
const created = await teamsClient.createTeam({
  displayName: "My team",
  description: "Optional",
  members: [{ userId: "owner-azure-ad-object-id", roles: ["owner"] }],
});
// Often HTTP 202: check `created.status === 202` and `operationLocation`
```

---

## API Reference

This reference lists the complete public API exported from `src/index.ts`.

### Authentication (Core)

#### Factories

- `createM365ClientCredentials(credentials)`  
  Creates an authentication instance with `GetAccessToken(scope?)` using a client secret.
- `createM365ClientCertificate(credentials)`  
  Creates an authentication instance with `GetAccessToken(scope?)` using certificate + private key.
- `getM365AuthenticationWithKeyVaultSigning(request)`  
  Builds an `M365Authentication` instance using the Key Vault signing flow.
- `getM365AuthenticationWithLocalCertificateSigning(request)`  
  Builds an `M365Authentication` instance using local certificate/key signing.

#### Token helpers

- `getM365AccessToken(credentials, scope?)`
- `getM365AccessTokenWithNodeSigning(request)`
- `getM365AccessTokenWithKeyVaultSigning(request)`
- `getM365AccessTokenWithLocalCertificateSigning(request)`

#### OOP authentication provider

- `M365AuthenticationProvider.buildWithKeyVaultSigning(request)`
- `M365AuthenticationProvider.buildWithLocalCertificateSigning(request)`

#### Local certificate helper

- `createM365LocalSigningCertificate(options)`  
  Creates a self-signed certificate for local development.

#### Key Vault helpers

- `getM365KeyVaultCertificate(request)`
- `getM365KeyVaultSecret(request)`
- `createM365KeyVaultJwtSigner(request)`

#### Constants / types

- `MS365Scopes.DEFAULT` = `https://graph.microsoft.com/.default`
- `MS365Scopes.KEY_VAULT` = `https://vault.azure.net/.default`
- Important types: `M365Authentication`, `M365AccessToken`, `MS365ClientCredentials`, `MS365CertificateCredentials`

---

### Teams

#### Factory

- `createTeamsClient(authentication): TeamsClient`

#### `TeamsClient` methods

- `getAccessToken(scope?)`  
  Returns an access token from the provided authentication provider.
- `getAllTeams(search?)`  
  `GET /v1.0/teams` (optional search parameter).
- `getTeamsBySearch(search?)`  
  Alias für `getAllTeams`.
- `getAllTeamsMetadata()`  
  Returns reduced team metadata.
- `getTeam(teamId)`  
  `GET /v1.0/teams/{teamId}`.
- `getTeamById(teamId)`  
  Alias für `getTeam`.
- `getAllTeamTemplates()`  
  `GET /v1.0/teamsTemplates`.
- `createTeam(input)`  
  `POST /v1.0/teams`; returns a team or a 202 provisioning result (`M365CreateTeamResult`).
- `deleteTeam(teamId)`  
  `DELETE /v1.0/teams/{teamId}`.

#### Teams types / guards

- Types: `M365Team`, `M365TeamMetadata`, `M365TeamTemplate`, `M365CreateTeamInput`, `M365CreateTeamResult`
- Guards: `isGraphTeamsResponse`, `isGraphTeamTemplatesResponse`, `isM365Team`

---

### SharePoint

#### Factory

- `createSharePointClient(authentication): SharePointClient`

#### Graph-based methods

- `getSharePointSite(siteId)`
- `getSharePointSiteById(siteId)` (Alias)
- `getAllSharePointSites(search?)`
- `getSharePointSitesBySearch(search?)` (Alias)
- `searchSharePointSitesWithSelect(search)`
- `getSharePointAllSitesMetadata()`
- `getSharePointLists(siteId)`  
  Deckt den Request `GET /v1.0/sites/{siteId}/lists?$select=id,displayName,list` ab.
- `getSharePointListColumns(siteId, listId)`  
  Deckt den Request `GET /v1.0/sites/{siteId}/lists/{listId}/columns` ab.
- `createSharePointListColumn(siteId, listId, column)`
- `createSharePointListColumns(siteId, listId, columns)`  
  Batch helper (sequential creation of multiple columns).

#### SharePoint REST view methods

- `getSharePointListViewXml({ siteWebUrl, listId, viewId })`
- `setSharePointListViewXml({ siteWebUrl, listId, viewId, viewXml })`
- `getSharePointDefaultViewByListTitle(siteWebUrl, listTitle)`
- `getSharePointDefaultViewByListId(siteWebUrl, listId)`
- `getSharePointListViewsByTitle(siteWebUrl, listTitle)`
- `getSharePointDefaultListViewXmlByTitle(siteWebUrl, listTitle, listId)`
- `getSharePointDefaultListViewXmlByListId(siteWebUrl, listId)`

#### View XML builder methods

- `buildSharePointViewXml(options)`  
  Builds new XML or replaces `ViewFields` in `baseViewXml`.
- `buildSharePointSetViewXmlPayload(options)`  
  Returns `{ viewXml }` directly for `SetViewXml`.
- `buildSharePointSetViewXmlPayloadFromViewXml(viewXml)`  
  Wrapper for existing XML strings.

#### SharePoint types / guards

- Site/List/Column types:  
  `M365SharePointSite`, `M365SharePointSiteMetadata`, `M365SharePointListInfo`, `M365SharePointColumn`
- Column create input types:  
  `M365SharePointColumnCreateInput` (union of `text`, `choice`, `boolean`, `number`, `dateTime`)
- View types:  
  `M365SharePointListViewXmlRequest`, `M365SharePointListViewInfo`, `M365SharePointSetViewXmlPayload`, `M365SharePointViewXmlBuilderOptions`
- Guards:  
  `isGraphSharePointSitesResponse`, `isGraphSharePointListsResponse`, `isGraphSharePointColumnsResponse`, `isM365SharePointSite`, `isM365SharePointListInfo`, `isM365SharePointColumn`

#### Example: create columns (multiple types)

```typescript
const createdColumns = await sharePointClient.createSharePointListColumns(
  "tenant365cloud.sharepoint.com,50362457-e8b6-45a2-8975-90f32004d97c,0bbcfbd6-25e3-4829-a029-a7fa0e562284",
  "157ec144-aef7-4130-88d8-3d5f07039de8",
  [
    {
      name: "projektnummer",
      displayName: "Projektnummer",
      description: "Projektnummer des Projekts",
      text: {},
    },
    {
      name: "aktenplan",
      displayName: "Aktenplan",
      description: "Aktenplan des Projekts",
      choice: {
        allowTextEntry: false,
        choices: ["AG-Angebot", "AG-Vertrag", "Dokumentation"],
        displayAs: "dropDownMenu",
      },
    },
    {
      name: "bereitzurarchivierung",
      displayName: "Bereit zur Archivierung",
      description: "Ist das Projekt bereit zur Archivierung?",
      boolean: {},
    },
  ],
);
```

---

### Entra (Microsoft Graph Directory)

#### Users

- Factory: `createEntraUsersClient(authentication)`
- Klasse: `EntraUsersClient`
- Methoden:
  - `getUser(userId)`
  - `getUserById(userId)` (Alias)
  - `getAllUsers(query?)`
  - `getUsersBySearch(search)` (`$search` wird korrekt formatiert/encoded)
  - `getAllUsersMetadata(query?)`

#### Groups

- Factory: `createEntraGroupsClient(authentication)`
- Klasse: `EntraGroupsClient`
- Methoden:
  - `getGroup(groupId)`
  - `getGroupById(groupId)` (Alias)
  - `getAllGroups(query?)`
  - `getGroupsBySearch(search)`
  - `getAllGroupsMetadata(query?)`

#### Applications

- Factory: `createEntraApplicationsClient(authentication)`
- Klasse: `EntraApplicationsClient`
- Methoden:
  - `getApplication(applicationId)`
  - `getApplicationById(applicationId)` (Alias)
  - `getAllApplications(query?)`
  - `getAllApplicationsMetadata(query?)`

#### Service Principals

- Factory: `createEntraServicePrincipalsClient(authentication)`
- Klasse: `EntraServicePrincipalsClient`
- Methoden:
  - `getServicePrincipal(servicePrincipalId)`
  - `getServicePrincipalById(servicePrincipalId)` (Alias)
  - `getAllServicePrincipals(query?)`
  - `getAllServicePrincipalsMetadata(query?)`

#### Directory Roles

- Factory: `createEntraDirectoryRolesClient(authentication)`
- Klasse: `EntraDirectoryRolesClient`
- Methoden:
  - `getDirectoryRole(roleId)`
  - `getDirectoryRoleById(roleId)` (Alias)
  - `getAllDirectoryRoles(query?)`
  - `getAllDirectoryRolesMetadata(query?)`
  - `getDirectoryRoleMembers(roleId)`

#### Entra helper

- `encodeGraphSearchParameter(search)`  
  Helper to normalize and encode Graph `$search` expressions.

---

## Token Response

All `GetAccessToken` methods return an `M365AccessToken`:

```typescript
type M365AccessToken = {
  token: string; // Bearer token
  expiresAt: Date; // Expiration time
};
```

---

## Debugging

Set the environment variable `TENANT365_MS_AUTH_DEBUG=1` to enable debug logging for authentication requests and responses.

```bash
TENANT365_MS_AUTH_DEBUG=1 node your-script.js
```

---

## Runtime Requirements

This library uses standard web platform APIs and is compatible with:

- **Node.js** 18+ (WebCrypto, Fetch, and TextEncoder are built-in)
- **Deno** and **Bun**
- Browser environments with WebCrypto support

Required globals: `crypto` (WebCrypto), `fetch`, `TextEncoder`, `atob`

---

## Build

```bash
pnpm install
pnpm build
```

## Test

```bash
pnpm test
```

Runs unit tests with Vitest.

---

## CI

A GitHub Actions workflow is included at `.github/workflows/pr-tests.yml`.
It runs type checks and unit tests for pull requests targeting `main`.
To block merges on failures, mark the workflow job as a required status check in branch protection settings.

Output is placed in `dist/`:

| File              | Format                  |
| ----------------- | ----------------------- |
| `dist/index.js`   | CommonJS                |
| `dist/index.mjs`  | ESM                     |
| `dist/index.d.ts` | TypeScript declarations |

---

## License

MIT – see [LICENSE](./LICENSE) for details.
