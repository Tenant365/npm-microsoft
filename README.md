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

### Authentication

#### `createM365ClientCredentials(credentials)`

Returns an object with a `GetAccessToken(scope?)` method that authenticates using a client secret.

| Parameter      | Type   | Description             |
| -------------- | ------ | ----------------------- |
| `tenantId`     | string | Azure AD tenant ID      |
| `clientId`     | string | Application (client) ID |
| `clientSecret` | string | Client secret value     |

#### `createM365ClientCertificate(credentials)`

Returns an object with a `GetAccessToken(scope?)` method that authenticates using a certificate and private key.

| Parameter     | Type                  | Description                                     |
| ------------- | --------------------- | ----------------------------------------------- |
| `tenantId`    | string                | Azure AD tenant ID                              |
| `clientId`    | string                | Application (client) ID                         |
| `privateKey`  | `CryptoKey` \| string | PEM-encoded private key or a `CryptoKey` object |
| `certificate` | string                | PEM-encoded X.509 certificate                   |
| `keyId`       | string (optional)     | Key ID for the JWT header                       |

Certificate authentication uses a JWT client assertion signed with RS256/RS384/RS512 (detected automatically from the key). The assertion includes the certificate thumbprint (`x5t` and `x5t#S256`).

#### `getM365AccessToken(credentials, scope?)`

Low-level function. Requests an access token using client credentials.

#### `getM365AccessTokenFromClientCertificate(credentials, scope?)`

Low-level function. Requests an access token using a certificate-based JWT assertion.

#### `getM365AccessTokenWithNodeSigning(request)`

Fetches the certificate from Key Vault, signs the JWT locally using the provided private key.

#### `getM365AccessTokenWithKeyVaultSigning(request)`

Fetches the certificate from Key Vault, signs the JWT remotely using the Key Vault sign API.

#### `getM365AccessTokenWithLocalCertificateSigning(request)`

Requests an access token using a local certificate + private key signing flow.

#### `getM365AuthenticationWithKeyVaultSigning(request)`

Builds an `M365Authentication` object using Key Vault-backed signing (certificate + signer). Useful when you want to pass authentication into service clients like `TeamsClient`.

#### `getM365AuthenticationWithLocalCertificateSigning(request)`

Builds an `M365Authentication` object from a local certificate/private key pair.

#### `M365AuthenticationProvider`

Object-oriented wrapper for authentication strategy selection:

- `buildWithKeyVaultSigning(request)`
- `buildWithLocalCertificateSigning(request)`

---

### Key Vault

#### `getM365KeyVaultCertificate(request)`

Fetches a certificate from Azure Key Vault.

```typescript
const cert = await getM365KeyVaultCertificate({
  vaultName: "my-vault",
  certificateName: "my-cert",
  authentication: createM365ClientCredentials({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
  }),
  // certificateVersion: "optional-version-id",
});

// cert.x509Pem       – PEM string
// cert.x509DerBase64 – base64-encoded DER bytes
```

#### `getM365KeyVaultSecret(request)`

Fetches a secret value from Azure Key Vault.

```typescript
const secret = await getM365KeyVaultSecret({
  vaultName: "my-vault",
  secretName: "my-secret",
  authentication: createM365ClientCredentials({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
  }),
});

// secret.value – the secret string
```

#### `createM365KeyVaultJwtSigner(request)`

Creates an async signer backed by Azure Key Vault. Useful when you need to sign arbitrary JWT payloads.

```typescript
const signer = createM365KeyVaultJwtSigner({
  vaultName: "my-vault",
  keyName: "my-key",
  authentication: createM365ClientCredentials({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
  }),
});

// signer.keyId  – Key Vault key ID
// signer.sign(signingInput, alg) – signs the input string using Key Vault
```

---

### Teams

#### `createTeamsClient(authentication)`

Creates a Teams client with helper methods:

- `getAccessToken()` – Requests a Graph token with `MS365Scopes.DEFAULT`
- `getAllTeams(search?)` – Calls `GET /v1.0/teams` (optional `$search` query, same pattern as SharePoint sites)
- `getTeamsBySearch(search?)` – Alias for `getAllTeams(search)`
- `getAllTeamsMetadata()` – Returns a reduced `M365TeamMetadata[]` from the teams list (same idea as `getSharePointAllSitesMetadata`)
- `getTeam(teamId)` – Calls `GET /v1.0/teams/{teamId}`
- `getTeamById(teamId)` – Alias for `getTeam(teamId)`
- `deleteTeam(teamId)` – Calls `DELETE /v1.0/teams/{teamId}`
- `getAllTeamTemplates()` – Calls `GET /v1.0/teamsTemplates` so you can choose a template id supported by your tenant
- `createTeam(input)` – Calls `POST /v1.0/teams` with `template@odata.bind` and `members` (Graph requirement). Returns the team resource or a **202** provisioning result with `operationLocation` / `contentLocation` headers.

List responses are validated with the exported helper `isGraphTeamsResponse` (expects `@odata.context` and `value`), matching `isGraphSharePointSitesResponse` for SharePoint.

#### Team creation: `404` / `CreateTeamFromTemplateRequest` / `Templates`

Graph accepts the HTTP request, but Microsoft Teams may still respond with **404 Not Found** and a message about **Templates** or **CreateTeamFromTemplateRequest**. That usually means the **template** or **tenant/Teams setup**, not a malformed JSON body from this library.

Checklist:

1. Call `getAllTeamTemplates()` and use a returned `id` as `templateId`, or set `templateOdataBind` to  
   `https://graph.microsoft.com/v1.0/teamsTemplates('<id>')` for that id.
2. Ensure **Microsoft Teams** is enabled for the tenant and the app has permissions such as `Team.Create` (and `User.Read.All` to resolve members).
3. Every `members[].userId` must be a real **Azure AD object id** of a user in that tenant (not `00000000-...`), and at least one **owner** is required.

---

### SharePoint

#### `createSharePointClient(authentication)`

Creates a SharePoint client with Graph and SharePoint REST helpers:

- `getAllSharePointSites(search?)` - `GET /v1.0/sites?search=...`
- `searchSharePointSitesWithSelect(search)` - `GET /v1.0/sites?search=<term>&$select=name,id,displayName,webUrl`
- `getSharePointLists(siteId)` - `GET /v1.0/sites/{siteId}/lists?$select=id,displayName,list`
- `getSharePointListColumns(siteId, listId)` - `GET /v1.0/sites/{siteId}/lists/{listId}/columns`
- `buildSharePointViewXml(options)` - builds View XML for list view updates
- `setSharePointListViewXml(request)` - `POST <siteWebUrl>/_api/Web/Lists(guid'{listId}')/Views(guid'{viewId}')/SetViewXml`

Example:

```typescript
const sharePointClient = createSharePointClient(auth);

const sites = await sharePointClient.searchSharePointSitesWithSelect("Controlx6Team");
const lists = await sharePointClient.getSharePointLists("hostname,siteCollectionId,siteId");
const columns = await sharePointClient.getSharePointListColumns(
  "hostname,siteCollectionId,siteId",
  "f3d9da8b-39d1-4567-9cec-996894b2ed78",
);

const viewXml = sharePointClient.buildSharePointViewXml({
  viewFields: ["Title", "Modified"],
  rowLimit: 100,
  whereClauseXml:
    "<Where><IsNotNull><FieldRef Name='Title' /></IsNotNull></Where>",
});

await sharePointClient.setSharePointListViewXml({
  siteWebUrl: "https://secnexdev.sharepoint.com/sites/Controlx11Team",
  listId: "f3d9da8b-39d1-4567-9cec-996894b2ed78",
  viewId: "8cf2f9a6-11a3-4eca-8a34-83b9fec192b2",
  viewXml,
});
```

---

### Entra (Microsoft Graph directory)

Modules live under `src/entra/` and are re-exported from the package root.

```typescript
import {
  createEntraUsersClient,
  createEntraGroupsClient,
  createEntraApplicationsClient,
  createEntraServicePrincipalsClient,
  createEntraDirectoryRolesClient,
} from "@tenant365/microsoft";

const users = createEntraUsersClient(auth);
const u = await users.getUser("user-object-id");
const allUsers = await users.getAllUsers();
// Search strings are normalized for Graph ($search): `displayName:Ann` becomes `"displayName:Ann"` and is URL-encoded.
const bySearch = await users.getUsersBySearch("displayName:Ann");

const groups = createEntraGroupsClient(auth);
const g = await groups.getGroup("group-object-id");

const apps = createEntraApplicationsClient(auth);
const appList = await apps.getAllApplications("$top=50");

const sp = createEntraServicePrincipalsClient(auth);
const principals = await sp.getAllServicePrincipals(
  "$filter=appId eq '00000000-0000-0000-0000-000000000000'",
);

const roles = createEntraDirectoryRolesClient(auth);
const roleList = await roles.getAllDirectoryRoles();
const members = await roles.getDirectoryRoleMembers("role-object-id");
```

| Factory | Client class | Graph resources |
| --- | --- | --- |
| `createEntraUsersClient` | `EntraUsersClient` | `/users` |
| `createEntraGroupsClient` | `EntraGroupsClient` | `/groups` |
| `createEntraApplicationsClient` | `EntraApplicationsClient` | `/applications` |
| `createEntraServicePrincipalsClient` | `EntraServicePrincipalsClient` | `/servicePrincipals` |
| `createEntraDirectoryRolesClient` | `EntraDirectoryRolesClient` | `/directoryRoles`, `/directoryRoles/{id}/members` |

List endpoints use the same `@odata.context` + `value` validation pattern as Teams and SharePoint. Grant the matching **Application** permissions (for example `User.Read.All`, `Group.Read.All`, `Application.Read.All`, `Directory.Read.All`) and admin consent as needed.

---

### Scopes

Predefined scope constants:

```typescript
import { MS365Scopes } from "@tenant365/microsoft";

MS365Scopes.DEFAULT; // "https://graph.microsoft.com/.default"
MS365Scopes.KEY_VAULT; // "https://vault.azure.net/.default"
```

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
