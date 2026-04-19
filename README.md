# Excel Online Client

A thin wrapper for Microsoft Graph API.

## Install

```console
npm install xls-graph-client
```

Also require.

```console
npm install @azure/indentity @microsoft/microsoft-graph-client
```

## Usage

```typescript
// Create credential
const credential = new ClientSecretCredential(
  "Tenant ID of your Tenant",
  "Client ID of your Entra ID application",
  "Client secret of your Entra ID application",
);

// Create Authentication provider
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ["openid", ".default"],
});

// Initialize client
const client = XlsGraphClient.initWithMiddleware({ authProvider });

// Open workbook by name
const workbook = await client.open("driveId", "itemIdOrName");

// Get range by address
const range = await workbook.worksheets("Sheet1").getRange("A1:X10");

// Get whole table (include header row)
const tableRange = await workbook.tables("Table1").getRange();

// Get header row
const tableHeader = await workbook.tables("Table1").getHeaderRowRange();

// Get data rows
const tableBody = await workbook.tables("Table1").getDataBodyRange();

for (const row of range) {
  for (const cell of row) {
    console.log(cell.value);
  }
}
```

## Samples

See samples in [the repository](https://github.com/sengokyu/typescript-excel-online-client/tree/main/samples).

## See also

- Dependent packages
  - https://www.npmjs.com/package/@microsoft/kiota-authentication-azure
  - https://www.npmjs.com/package/@microsoft/msgraph-sdk
  - https://www.npmjs.com/package/@microsoft/msgraph-sdk-drives
- Document
  - [Microsoft Graph REST API v1.0 endpoint reference](https://learn.microsoft.com/en-us/graph/api/overview)
  - [Working with Excel in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel)
