import { ExternalConnectors } from '@microsoft/microsoft-graph-types';

export const config = {
  connection: {
    id: '{{connectionId}}',
    name: '{{connectorName}}',
    description: '{{connectorDescription}}',
    searchSettings: {
      searchResultTemplates: [
        {
          id: '{{connectionId}}',
          priority: 1,
          layout: {}
        }
      ]
    },
    // https://learn.microsoft.com/graph/connecting-external-content-manage-schema
    schema: {
      baseType: 'microsoft.graph.externalItem',
      properties: []
    }
  } as ExternalConnectors.ExternalConnection
};