export const config = {
  tenantId: process.env.XOC_TENANT_ID ?? "consumers",
  clientId: process.env.XOC_CLIENT_ID ?? "f892dc14-e9be-4ec9-822d-5a0debb485c1",
  driveId: process.env.XOC_DRIVE_ID ?? "D5A9D3A9B9695ED2",
  itemId: process.env.XOC_ITEM_ID ?? "D5A9D3A9B9695ED2!1314",
  scopes: [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Files.Read",
  ],
};
