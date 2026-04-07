import selfsigned from "selfsigned";

export interface M365LocalSigningCertificateOptions {
  commonName?: string;
  organizationName?: string;
  countryName?: string;
  daysValid?: number;
  keySize?: 2048 | 3072 | 4096;
}

export interface M365LocalSigningCertificate {
  privateKeyPem: string;
  certificatePem: string;
  publicKeyPem: string;
  fingerprintSha1?: string;
}

/**
 * Creates a local self-signed certificate bundle for local JWT client assertion signing.
 * This is intended for development/test or controlled environments.
 */
export const createM365LocalSigningCertificate = async (
  options: M365LocalSigningCertificateOptions = {},
): Promise<M365LocalSigningCertificate> => {
  const attrs = [
    { name: "commonName", value: options.commonName ?? "tenant365.local" },
    {
      name: "organizationName",
      value: options.organizationName ?? "Tenant365",
    },
    { name: "countryName", value: options.countryName ?? "DE" },
  ];

  const generated: any = await (selfsigned as any).generate(attrs, {
    notAfter: options.daysValid ?? 365,
    keySize: options.keySize ?? 2048,
    algorithm: "sha256",
  });

  return {
    privateKeyPem: generated.private,
    certificatePem: generated.cert,
    publicKeyPem: generated.public,
    fingerprintSha1: generated.fingerprint,
  };
};
