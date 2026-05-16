import { XlsGraphClient } from "xls-graph-client";
import { config } from "./config.js";
import { BrowserInteractiveCredential } from "./credential.js";

const sheetName = "Sheet1";
const rangeAddress = "A1:C3";

const authProvider = new BrowserInteractiveCredential(config);
const client = XlsGraphClient.createInstance(authProvider);

const workbook = await client.open(config.driveId, config.itemId);
const range = await workbook.worksheets(sheetName).getRange(rangeAddress);

range.forEach((row, rowIndex) => {
  row.forEach((cell, cellIndex) => {
    console.log(
      `Value at ${String.fromCharCode(65 + cellIndex)}${rowIndex + 1}: ${cell.value}`,
    );
  });
});
