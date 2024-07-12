import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import fs from 'fs';
import { config } from './config.js';
import { client } from './graphClient.js';

async function createConnection(): Promise<void> {
  console.log('Creating connection...');

  const { id, name, description, activitySettings, searchSettings } = config.connection;
  // update item resolvers
  activitySettings?.urlToItemResolvers?.forEach((r: ExternalConnectors.ItemIdResolver) => {
    (r as any)['@odata.type'] = '#microsoft.graph.externalConnectors.itemIdResolver';
  })
  const adaptiveCard = fs.readFileSync('./resultLayout.json', 'utf8');
  searchSettings!.searchResultTemplates![0].layout = JSON.parse(adaptiveCard);

  await client
    .api('/external/connections')
    .post({
      id,
      name,
      description,
      activitySettings,
      searchSettings
    });

  console.log('Connection created');
}

async function createSchema(): Promise<void> {
  console.log('Creating schema...');

  const { id, schema } = config.connection;
  try {
    const res = await client
      .api(`/external/connections/${id}/schema`)
      .header('content-type', 'application/json')
      .patch(schema) as ExternalConnectors.ConnectionOperation;

    const status = res.status;
    if (status === 'completed') {
      console.log('Schema created');
    }
    else {
      console.error(`Schema creation failed: ${res.error?.message}`);
    }
  }
  catch (e) {
    console.error(e);
  }
}

async function main() {
  try {
    await createConnection();
    await createSchema();
  }
  catch (e) {
    console.error(e);
  }
}

main();