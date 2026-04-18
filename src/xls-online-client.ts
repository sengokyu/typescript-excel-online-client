import type { AuthenticationProvider } from "@microsoft/kiota-abstractions";
import {
  createGraphServiceClient,
  GraphRequestAdapter,
  GraphServiceClient,
} from "@microsoft/msgraph-sdk";
import { Workbook } from "./workbook.js";

/**
 * A thin wrapper for MS graph api
 */
export class XlsOnlineClient {
  /**
   * cTor
   */
  private constructor(private readonly client: GraphServiceClient) {}

  /**
   * Create a instance.
   * @public
   * @static
   * @param authenticationProvider
   * @returns
   */
  static createInstance(
    authenticationProvider: AuthenticationProvider,
  ): XlsOnlineClient {
    const requestAdapter = new GraphRequestAdapter(authenticationProvider);
    const client = createGraphServiceClient(requestAdapter);

    return new XlsOnlineClient(client);
  }

  /**
   * GraphServiceClient
   */
  public get underlyingObject(): GraphServiceClient {
    return this.client;
  }

  /**
   * Open a workbook
   * @param driveId The drive ID (e.g. the ID of the user's OneDrive)
   * @param itemId The drive item ID of the Excel file
   */
  public openWorkbook(driveId: string, itemId: string): Promise<Workbook> {
    return Workbook.createInstance(this.client, driveId, itemId);
  }
}
