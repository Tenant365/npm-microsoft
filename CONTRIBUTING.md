# Contributing

## Branching model: release branches

This repository uses **release branches** for version stabilization and publishing, instead of shipping directly from short-lived feature branches.

### Long-lived branches

| Branch    | Purpose                                                                                          |
| --------- | ------------------------------------------------------------------------------------------------ |
| `main`    | Production-ready code; tagged releases are cut from here (or from a release branch merged here). |
| `develop` | Integration branch for ongoing work (optional; use if your team prefers Git Flow).               |

### Branch types

1. **`feature/<short-description>`**
   - Branch from: `develop` (or `main` if you do not use `develop`).
   - Open a **pull request** back to `develop` / `main`.
   - Used for new functionality.

2. **`release/v<major>.<minor>.<patch>`** (example: `release/v0.1.0`)
   - Branch from: `develop` or `main` when you start a release.
   - Only **bugfixes, docs, and version bumps** (e.g. `package.json`).
   - Open a **pull request** to `main` when the release is ready.
   - After merge: tag `v0.1.0` on `main` and publish to npm as per your process.

3. **`hotfix/<short-description>`** (optional)
   - Branch from: `main` for urgent production fixes.
   - Merge to `main` and back-port to `develop` if applicable.

### Typical release flow

```text
develop (or main)
    │
    ├── feature/foo ──► PR ──► develop
    │
    └── create release/v1.2.0 from develop
              │
              ├── fix only on release/v1.2.0
              │
              └── PR release/v1.2.0 ──► main ──► tag v1.2.0 ──► npm publish
```

### CI

Pull requests targeting **`main`** or **`release/**`** run typecheck and tests (see `.github/workflows/pr-tests.yml`).

Configure GitHub **branch protection** so merges to `main` and `release/*` require that check to pass.

### Commits

Use clear commit messages (e.g. [Conventional Commits](https://www.conventionalcommits.org/)): `feat:`, `fix:`, `docs:`, `chore:`.
