import {
  getM365AuthenticationWithKeyVaultSigning,
  MS365Scopes,
} from "./core/auth";
import { createTeamsClient } from "./teams";

const run = async () => {
  const authentication = await getM365AuthenticationWithKeyVaultSigning({
    tenantId: process.env.M365_TENANT_ID,
    clientId: process.env.M365_CLIENT_ID,
    clientSecret: process.env.M365_CLIENT_SECRET,
    keyVaultName: process.env.M365_KEY_VAULT_NAME,
    certificateName: process.env.M365_CERTIFICATE_NAME,
    keyName: process.env.M365_KEY_NAME,
    keyVaultTenantId: process.env.M365_KEY_VAULT_TENANT_ID,
  });

  const teamsClient = createTeamsClient(authentication);

  const accessToken = await teamsClient.getAccessToken(MS365Scopes.DEFAULT);
  const teams = await teamsClient.getAllTeamsWithAccessToken(accessToken);

  console.log("Teams:", teams);
};

void run();
