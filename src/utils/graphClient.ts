import { Client } from '@microsoft/microsoft-graph-client'
import { ClientSecretCredential } from "@azure/identity"
import { config } from '../utils/config'

const credential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);

export const client = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: async () => {
      const token = await credential.getToken("https://graph.microsoft.com/.default");
      return token.token;
    }
  }
});