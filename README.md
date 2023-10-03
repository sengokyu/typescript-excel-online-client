# Excel Online Client

A thin wrapper for Microsoft Graph API.

## Install

```console
npm install excel-graph-client
```

Also require.

```console
npm install @azure/indentity @microsoft/microsoft-graph-client
```
## Usage

```typescript
// Create credential
const credential = new ClientSecretCredential(
  'Tenant ID of your Azure Active Directory',
  'Client ID of your Azure Active Directory application',
  'Client secret of your Azure Active Directory application'
);

// Create Authentication provider
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ["openid", ".default"],
});

// Initialize client
const client = ExcelOnlineClient.initWithMiddleware({ authProvider });

// Open workbook by name
const workbook = await client.open("me/drive/root:/Excel/ExcelGraphSample.xlsx");

// Get worksheet by name
const worksheet = await workbook.getWorksheet("Sheet1");

// Get range by address
const values = await worksheet.getRange("A1:X10");

for (const row of values) {
    for (const column of row) {
        console.log(column);
    }
}
```
