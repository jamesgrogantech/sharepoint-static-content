// eslint-disable-next-line @typescript-eslint/no-var-requires
const { ConfidentialClientApplication } = require('@azure/msal-node');

const clientId = process.env.CLIENT_ID;
const tenantId = process.env.TENANT_ID;
const username = process.env.SHAREPOINT_USERNAME;
const password = process.env.SHAREPOINT_PASSWORD;
const clientSecret = process.env.CLIENT_SECRET;

const scopes = ['https://graph.microsoft.com/Sites.Read.All'];

export async function getAccessToken() {
  const cca = new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  });

  const result = await cca.acquireTokenByUsernamePassword({
    username,
    password,
    scopes,
  });

  return result.accessToken;
}
