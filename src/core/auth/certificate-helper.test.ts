import { describe, expect, it } from "vitest";
import { createM365LocalSigningCertificate } from "./certificate-helper";

describe("createM365LocalSigningCertificate", () => {
  it("creates PEM formatted key/certificate bundle", async () => {
    const result = await createM365LocalSigningCertificate({
      commonName: "example.local",
      daysValid: 30,
    });

    expect(result.privateKeyPem).toContain("BEGIN PRIVATE KEY");
    expect(result.certificatePem).toContain("BEGIN CERTIFICATE");
    expect(result.publicKeyPem).toContain("BEGIN PUBLIC KEY");
  });
});
