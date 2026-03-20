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

  await teamsClient.getAccessToken();
  const teams = await teamsClient.getAllTeams();

  const team = await teamsClient.getTeam(
    "00000000-0000-0000-0000-000000000000",
  );

  console.log("Teams:", teams);
  console.log("Team:", team);
};

void run();
