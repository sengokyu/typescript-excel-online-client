import { setLogLevel } from "@azure/logger";
import { AzureIdentityAuthenticationProvider } from "@microsoft/kiota-authentication-azure";
import { ExcelOnlineClient } from "excel-graph-client";
import { config } from "./config.js";
import { BrowserInteractiveCredential } from "./credential.js";

setLogLevel("error");

const sheetName = "Sheet1";
const tableName = "Table1";

const credential = new BrowserInteractiveCredential(config);
const authProvider = new AzureIdentityAuthenticationProvider(
  credential,
  config.scopes,
  { tenantId: config.tenantId, enableCae: false },
);

const client = ExcelOnlineClient.createInstance(authProvider);

const workbook = await client.openWorkbook(config.driveId, config.itemId);
const worksheet = await workbook.getWorksheet(sheetName);
const table = await worksheet.getTable(tableName);
const range = await table.getRange();

for (const row of range) {
  for (const cell of row) {
    console.log(cell.value);
  }
}
