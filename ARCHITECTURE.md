# Architecture

## Goals

- Provide reusable Microsoft Graph clients for multiple domains.
- Keep authentication strategies explicit (client secret, local certificate signing, Key Vault signing).
- Use consistent OOP client patterns across modules.

## Source layout

- `src/core/auth`
  - Authentication primitives and token acquisition.
  - Exposes strategy-specific builders and OOP provider wrapper.
- `src/sharepoint`
  - SharePoint Graph client split into:
    - `types.ts` (domain contracts)
    - `guards.ts` (runtime shape validation)
    - `client.ts` (OOP client implementation)
    - `index.ts` (module barrel)
- `src/teams`
  - Teams Graph client split into:
    - `types.ts` (domain contracts)
    - `guards.ts` (runtime shape validation)
    - `client.ts` (OOP client implementation)
    - `index.ts` (module barrel)
- `src/entra`
  - Entra directory modules:
    - `users`, `groups`, `applications`, `service-principals`, `directory-roles`
  - Shared helpers in `src/entra/internal`.
- `src/common/graph`
  - Shared Graph transport and HTTP behavior (`requestM365Graph`).

## Design conventions

### 1) Client class per domain

Each domain exports a client class and a factory function:

- `TeamsClient` + `createTeamsClient`
- `SharePointClient` + `createSharePointClient`
- `EntraUsersClient` + `createEntraUsersClient` (and equivalents)

Clients receive `M365Authentication` in the constructor and resolve tokens internally.

### 2) Runtime guards for Graph responses

TypeScript types are compile-time only, so responses are validated with runtime guards before returning typed values.

Examples:

- `isGraphTeamsResponse`
- `isGraphSharePointSitesResponse`
- `isGraphEntraUsersResponse`

### 3) Explicit auth strategy names

The code uses distinct function names per signing strategy:

- Key Vault signing:
  - `getM365AuthenticationWithKeyVaultSigning`
  - `getM365AccessTokenWithKeyVaultSigning`
- Local certificate signing:
  - `getM365AuthenticationWithLocalCertificateSigning`
  - `getM365AccessTokenWithLocalCertificateSigning`

This avoids ambiguity and keeps callsites readable.

### 4) OOP auth provider wrapper

`M365AuthenticationProvider` offers strategy methods:

- `buildWithKeyVaultSigning(...)`
- `buildWithLocalCertificateSigning(...)`

Use this when you want object-oriented orchestration in app code.

### 5) Shared transport layer (enterprise style)

All modules should use the same Graph transport helper:

- `requestM365Graph(path, accessToken, options)`

Benefits:

- one place for HTTP conventions
- one place for JSON/error normalization
- consistent error messages across domains
- easier telemetry/retry extension later

### 6) Shared base class for domain clients

`M365GraphClientBase` centralizes:

- token resolution (`getAccessToken`)
- Graph request dispatch (`graphRequest`)

Domain clients (Teams and SharePoint now, Entra incrementally) extend this base class to reduce duplication and keep behavior uniform.

## Request flow

1. Resolve access token from `M365Authentication`.
2. Execute Graph request.
3. Parse JSON (with text fallback for non-JSON errors).
4. Validate response shape via guard.
5. Return domain type or throw a descriptive error.

## Testing strategy

- Unit tests with Vitest per domain module.
- `fetch` mocked for deterministic behavior.
- Focus on:
  - token delegation
  - request contract validation
  - response-shape validation
  - failure-path error messaging

## CI and release model

- CI workflow: `.github/workflows/pr-tests.yml`
  - runs `tsc --noEmit` and `pnpm test`.
- Release flow is documented in `CONTRIBUTING.md`.
  - feature branches -> develop/main
  - release branches `release/vX.Y.Z` -> main

## Extension guidelines

When adding a new Graph domain:

1. Create folder `src/<domain>`.
2. Add `types.ts`, `guards.ts`, `client.ts`, `index.ts`.
3. Export via `src/index.ts`.
4. Add tests under the domain folder.
5. Document in README + this architecture doc.
