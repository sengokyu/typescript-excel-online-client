# Excel Online Client

A thin wrapper for Microsoft Graph API.

## Install

```console
npm install xls-graph-client
```

Also require.

```console
npm install @microsoft/msgraph-sdk @microsoft/msgraph-sdk-drives
```

## Usage

```typescript
import type { AuthenticationProvider } from "@microsoft/kiota-abstractions";

// Implement AuthenticationProvider with any OAuth2 library (e.g. openid-client)
const authProvider: AuthenticationProvider = /* */;

// Initialize client
const client = XlsGraphClient.createInstance(authProvider);

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
  - https://www.npmjs.com/package/@microsoft/msgraph-sdk
  - https://www.npmjs.com/package/@microsoft/msgraph-sdk-drives
- Authentication
  - https://www.npmjs.com/package/openid-client
- Document
  - [Microsoft Graph REST API v1.0 endpoint reference](https://learn.microsoft.com/en-us/graph/api/overview)
  - [Working with Excel in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel)
