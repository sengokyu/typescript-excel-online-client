import * as graph from "@microsoft/microsoft-graph-client";
import { Workbook } from "./workbook";

/**
 * A thin wrapper of Client
 */
export class ExcelOnlineClient {
  /**
   * cTor
   */
  private constructor(private readonly client: graph.Client) {}

  /**
   * Create a instance.
   * @public
   * @static
   * @param options
   * @returns
   */
  static init(options: graph.Options): ExcelOnlineClient {
    return new ExcelOnlineClient(graph.Client.init(options));
  }

  /**
   * Create a instance.
   * @public
   * @static
   * @param clientOptions
   * @returns
   */
  static initWithMiddleware(
    clientOptions: graph.ClientOptions
  ): ExcelOnlineClient {
    return new ExcelOnlineClient(
      graph.Client.initWithMiddleware(clientOptions)
    );
  }

  /**
   * Open a workbook
   */
  public open(workbookPath: string): Promise<Workbook> {
    return Workbook.init(this.client, workbookPath);
  }
}
