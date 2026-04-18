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
const client = XlsOnlineClient.initWithMiddleware({ authProvider });

// Open workbook by name
const workbook = await client.open(
  "me/drive/root:/Excel/ExcelGraphSample.xlsx",
);

// Get worksheet by name
const worksheet = await workbook.getWorksheet("Sheet1");

// Get range by address
const range = await worksheet.getRange("A1:X10");

for (const row of range) {
  for (const cell of row) {
    console.log(cell.value);
  }
}

// Get table by name
const table = await worksheet.getTable("Table1");

// Get whole table (include header row)
const tableRange = await table.getRange();
// Get header row
const tableHeader = await table.getHeaderRowRange();
// Get data rows
const tableBody = await table.getDataBodyRange();
```

## See also

- Dependent packages
  - https://www.npmjs.com/package/@microsoft/kiota-authentication-azure
  - https://www.npmjs.com/package/@microsoft/msgraph-sdk
  - https://www.npmjs.com/package/@microsoft/msgraph-sdk-drives
- Document
  - [Microsoft Graph REST API v1.0 endpoint reference](https://learn.microsoft.com/en-us/graph/api/overview)
  - [Working with Excel in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel)
