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

import { getM365AuthenticationWithKeyVaultSigning } from "./access-token";

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
