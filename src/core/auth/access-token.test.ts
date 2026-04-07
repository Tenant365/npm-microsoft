import { describe, expect, it, vi, beforeEach } from "vitest";

const {
  createM365ClientCredentialsMock,
  createM365ClientCertificateMock,
  getM365KeyVaultCertificateMock,
  createM365KeyVaultJwtSignerMock,
} = vi.hoisted(() => ({
  createM365ClientCredentialsMock: vi.fn(),
  createM365ClientCertificateMock: vi.fn(),
  getM365KeyVaultCertificateMock: vi.fn(),
  createM365KeyVaultJwtSignerMock: vi.fn(),
}));

vi.mock("./auth", () => ({
  createM365ClientCredentials: createM365ClientCredentialsMock,
  createM365ClientCertificate: createM365ClientCertificateMock,
}));

vi.mock("./keyvault", () => ({
  getM365KeyVaultCertificate: getM365KeyVaultCertificateMock,
  createM365KeyVaultJwtSigner: createM365KeyVaultJwtSignerMock,
}));

import {
  getM365AuthenticationWithKeyVaultSigning,
  getM365AuthenticationWithLocalCertificateSigning,
  M365AuthenticationProvider,
} from "./access-token";

describe("getM365AuthenticationWithKeyVaultSigning", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("builds key vault based authentication flow", async () => {
    const keyVaultAuth = { GetAccessToken: vi.fn() };
    const signer = { keyId: "kid-1", sign: vi.fn() };
    const auth = { GetAccessToken: vi.fn() };

    createM365ClientCredentialsMock.mockReturnValue(keyVaultAuth);
    getM365KeyVaultCertificateMock.mockResolvedValue({
      x509Pem: "-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----",
    });
    createM365KeyVaultJwtSignerMock.mockReturnValue(signer);
    createM365ClientCertificateMock.mockReturnValue(auth);

    const result = await getM365AuthenticationWithKeyVaultSigning({
      tenantId: "tenant-a",
      clientId: "client-a",
      clientSecret: "secret-a",
      keyVaultName: "kv-a",
      certificateName: "cert-a",
      keyName: "key-a",
      keyVaultTenantId: "tenant-kv",
      keyVaultClientId: "client-kv",
      keyVaultClientSecret: "secret-kv",
      certificateVersion: "cert-v1",
      keyVersion: "key-v1",
    });

    expect(createM365ClientCredentialsMock).toHaveBeenCalledWith({
      tenantId: "tenant-kv",
      clientId: "client-kv",
      clientSecret: "secret-kv",
    });

    expect(getM365KeyVaultCertificateMock).toHaveBeenCalledWith({
      vaultName: "kv-a",
      certificateName: "cert-a",
      certificateVersion: "cert-v1",
      authentication: keyVaultAuth,
    });

    expect(createM365KeyVaultJwtSignerMock).toHaveBeenCalledWith({
      vaultName: "kv-a",
      keyName: "key-a",
      keyVersion: "key-v1",
      authentication: keyVaultAuth,
    });

    expect(createM365ClientCertificateMock).toHaveBeenCalledWith({
      tenantId: "tenant-a",
      clientId: "client-a",
      certificate:
        "-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----",
      keyVaultSigner: signer,
      keyId: "kid-1",
    });

    expect(result).toBe(auth);
  });
});

describe("local certificate signing auth", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("builds authentication from local certificate and private key", async () => {
    const auth = { GetAccessToken: vi.fn() };
    createM365ClientCertificateMock.mockReturnValue(auth);

    const result = await getM365AuthenticationWithLocalCertificateSigning({
      tenantId: "tenant-local",
      clientId: "client-local",
      privateKey: "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----",
      certificate: "-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----",
    });

    expect(createM365ClientCertificateMock).toHaveBeenCalledWith({
      tenantId: "tenant-local",
      clientId: "client-local",
      privateKey: "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----",
      certificate: "-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----",
    });
    expect(result).toBe(auth);
  });

  it("provider class delegates to local and keyvault builders", async () => {
    const provider = new M365AuthenticationProvider();
    const localAuth = { GetAccessToken: vi.fn() };
    const kvAuth = { GetAccessToken: vi.fn() };

    createM365ClientCertificateMock.mockReturnValueOnce(localAuth);
    createM365ClientCredentialsMock.mockReturnValue({ GetAccessToken: vi.fn() });
    getM365KeyVaultCertificateMock.mockResolvedValue({ x509Pem: "cert" });
    createM365KeyVaultJwtSignerMock.mockReturnValue({
      keyId: "kid-provider",
      sign: vi.fn(),
    });
    createM365ClientCertificateMock.mockReturnValueOnce(kvAuth);

    const localResult = await provider.buildWithLocalCertificateSigning({
      tenantId: "tenant-local",
      clientId: "client-local",
      privateKey: "pk",
      certificate: "cert",
    });
    const kvResult = await provider.buildWithKeyVaultSigning({
      tenantId: "tenant-kv",
      clientId: "client-kv",
      clientSecret: "secret-kv",
      keyVaultName: "vault",
      certificateName: "cert",
      keyName: "key",
    });

    expect(localResult).toBe(localAuth);
    expect(kvResult).toBe(kvAuth);
  });
});
