import { setLogLevel } from "@azure/logger";
import { AzureIdentityAuthenticationProvider } from "@microsoft/kiota-authentication-azure";
import { XlsGraphClient } from "xls-graph-client";
import { config } from "./config.js";
import { BrowserInteractiveCredential } from "./credential.js";

setLogLevel("error");

const tableName = "Table1";

const credential = new BrowserInteractiveCredential(config);
const authProvider = new AzureIdentityAuthenticationProvider(
  credential,
  config.scopes,
  { tenantId: config.tenantId, enableCae: false },
);

const client = XlsGraphClient.createInstance(authProvider);

const workbook = await client.open(config.driveId, config.itemId);
const range = await workbook.tables(tableName).getRange();

for (const row of range) {
  for (const cell of row) {
    console.log(cell.value);
  }
}
