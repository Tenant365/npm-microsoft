# @tenant365/microsoft

Microsoft APIs – Powered by Tenant365

A TypeScript library for Microsoft 365 and Azure authentication. Supports client credentials, certificate-based authentication, and Azure Key Vault-backed signing.

---

## Features

- **Client Credentials** – Authenticate with a client secret
- **Certificate Authentication** – Sign JWT assertions locally with a private key
- **Key Vault Signing** – Sign JWT assertions remotely via Azure Key Vault (private key never leaves the vault)
- **Key Vault Utilities** – Fetch certificates and secrets from Azure Key Vault
- **Dual module output** – CommonJS and ESM, with TypeScript declarations

---

## Installation

```bash
npm install @tenant365/microsoft
# or
pnpm add @tenant365/microsoft
```

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

### Key Vault Signing

The private key never leaves Azure Key Vault. The library fetches the certificate and delegates signing to the Key Vault sign API.

```typescript
import { getM365AccessTokenWithKeyVaultSigning, MS365Scopes } from "@tenant365/microsoft";

const { token, expiresAt } = await getM365AccessTokenWithKeyVaultSigning({
  tenantId: "your-tenant-id",
  clientId: "your-client-id",
  keyVaultCredentials: {
    tenantId: "your-tenant-id",
    clientId: "your-client-id",
    clientSecret: "your-keyvault-client-secret",
  },
  keyVaultName: "your-vault-name",
  certificateName: "your-certificate-name",
  keyName: "your-key-name",
  scope: MS365Scopes.DEFAULT,
});
```

---

## API Reference

### Authentication

#### `createM365ClientCredentials(credentials)`

Returns an object with a `GetAccessToken(scope?)` method that authenticates using a client secret.

| Parameter      | Type   | Description                          |
|----------------|--------|--------------------------------------|
| `tenantId`     | string | Azure AD tenant ID                   |
| `clientId`     | string | Application (client) ID              |
| `clientSecret` | string | Client secret value                  |

#### `createM365ClientCertificate(credentials)`

Returns an object with a `GetAccessToken(scope?)` method that authenticates using a certificate and private key.

| Parameter    | Type                   | Description                                         |
|--------------|------------------------|-----------------------------------------------------|
| `tenantId`   | string                 | Azure AD tenant ID                                  |
| `clientId`   | string                 | Application (client) ID                             |
| `privateKey` | `CryptoKey` \| string  | PEM-encoded private key or a `CryptoKey` object     |
| `certificate`| string                 | PEM-encoded X.509 certificate                       |
| `keyId`      | string (optional)      | Key ID for the JWT header                           |

Certificate authentication uses a JWT client assertion signed with RS256/RS384/RS512 (detected automatically from the key). The assertion includes the certificate thumbprint (`x5t` and `x5t#S256`).

#### `getM365AccessToken(credentials, scope?)`

Low-level function. Requests an access token using client credentials.

#### `getM365AccessTokenFromClientCertificate(credentials, scope?)`

Low-level function. Requests an access token using a certificate-based JWT assertion.

#### `getM365AccessTokenWithNodeSigning(request)`

Fetches the certificate from Key Vault, signs the JWT locally using the provided private key.

#### `getM365AccessTokenWithKeyVaultSigning(request)`

Fetches the certificate from Key Vault, signs the JWT remotely using the Key Vault sign API.

---

### Key Vault

#### `getM365KeyVaultCertificate(request)`

Fetches a certificate from Azure Key Vault.

```typescript
const cert = await getM365KeyVaultCertificate({
  credentials: { tenantId, clientId, clientSecret },
  keyVaultName: "my-vault",
  certificateName: "my-cert",
  // version: "optional-version-id",
});

// cert.certificatePem  – PEM string
// cert.certificateDer  – base64-encoded DER bytes
```

#### `getM365KeyVaultSecret(request)`

Fetches a secret value from Azure Key Vault.

```typescript
const secret = await getM365KeyVaultSecret({
  credentials: { tenantId, clientId, clientSecret },
  keyVaultName: "my-vault",
  secretName: "my-secret",
});

// secret.value – the secret string
```

#### `createM365KeyVaultJwtSigner(request)`

Creates an async signer backed by Azure Key Vault. Useful when you need to sign arbitrary JWT payloads.

```typescript
const signer = await createM365KeyVaultJwtSigner({
  credentials: { tenantId, clientId, clientSecret },
  keyVaultName: "my-vault",
  keyName: "my-key",
});

// signer.keyId  – Key Vault key ID
// signer.sign(signingInput, alg) – signs the input string using Key Vault
```

---

### Scopes

Predefined scope constants:

```typescript
import { MS365Scopes } from "@tenant365/microsoft";

MS365Scopes.DEFAULT    // "https://graph.microsoft.com/.default"
MS365Scopes.KEY_VAULT  // "https://vault.azure.net/.default"
```

---

## Token Response

All `GetAccessToken` methods return an `M365AccessToken`:

```typescript
type M365AccessToken = {
  token: string;   // Bearer token
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

Output is placed in `dist/`:

| File              | Format      |
|-------------------|-------------|
| `dist/index.js`   | CommonJS    |
| `dist/index.mjs`  | ESM         |
| `dist/index.d.ts` | TypeScript declarations |

---

## License

MIT – see [LICENSE](./LICENSE) for details.
