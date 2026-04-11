import { GraphServiceClient } from "@microsoft/msgraph-sdk";
import { Cell } from "./cell";

/**
 *
 */
export class Worksheet {
  private constructor(
    private readonly client: GraphServiceClient,
    private readonly worksheetPath: string,
  ) {
    // TODO
  }

  /**
   *
   * @param client
   * @param worksheetPath
   */
  static async init(
    client: GraphServiceClient,
    worksheetPath: string,
  ): Promise<Worksheet> {
    // Check existence.

    // TODO

    return new Worksheet(client, worksheetPath);
  }

  /**
   * Get values of the range
   * @param address
   * @returns
   */
  async getRange(address: string): Promise<Cell[][]> {
    // TODO
  }
}
