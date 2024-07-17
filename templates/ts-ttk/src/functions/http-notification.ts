import { app, HttpRequest, InvocationContext } from "@azure/functions";
import { getEntraJwksUri, TokenValidator, ValidateTokenOptions } from "jwt-validate";
import { config } from "../common/config";
import { ConnectionMessage } from "../common/ConnectionMessage";
import { getQueueClient } from "../common/queueClient";
import { streamToJson } from "../common/utils";

enum TargetConnectorState {
  Enabled = 'enabled',
  Disabled = 'disabled',
}

app.http('notification', {
  methods: ['POST'],
  handler: async (request: HttpRequest, console: InvocationContext) => {

    (async () => {
      const body = await streamToJson(request.body);
      console.log('Received notification');
      console.log(JSON.stringify(body, null, 2));

      const {
        aadAppTenantId: tenantId,
        aadAppClientId: clientId
      } = config;

      const token = body?.validationTokens[0];
      console.log(`Validating token: ${token}, tenantId: ${tenantId}, clientId: ${clientId}...`);
      const entraJwksUri = await getEntraJwksUri(tenantId);
      const validator = new TokenValidator({
        jwksUri: entraJwksUri
      });
      const options: ValidateTokenOptions = {
        audience: clientId,
        issuer: `https://login.microsoftonline.com/${tenantId}/v2.0`,
        idtyp: 'app',
        ver: '2.0'
      };
      await validator.validateToken(token, options);
      console.log('Token validated');

      const changeDetails = body?.value[0]?.resourceData;
      const targetConnectorState = changeDetails?.state;

      const message: ConnectionMessage = {
        connectorId: changeDetails?.id,
        connectorTicket: changeDetails?.connectorsTicket
      }

      if (targetConnectorState === TargetConnectorState.Enabled) {
        message.action = 'create';
      }
      else if (targetConnectorState === TargetConnectorState.Disabled) {
        message.action = 'delete';
      }

      if (!message.action) {
        console.error('Invalid action');
        return;
      }

      console.log(JSON.stringify(message, null, 2));

      const queueClient = await getQueueClient('queue-connection');
      const messageString = btoa(JSON.stringify(message));
      console.log('Sending message to queue queue-connection: ${message}');
      // must base64 encode
      await queueClient.sendMessage(messageString);
      console.log('Message sent');
    })();

    return {
      status: 202
    }
  }
})