import { ExcelOnlineClient } from "./excel-online-client";
import * as graph from "@microsoft/microsoft-graph-client";
import * as identity from "@azure/identity";

jest.mock("@microsoft/microsoft-graph-client");

describe("ExcelOnlineClient", () => {
  beforeEach(() => {
    (graph.Client as unknown as jest.Mock).mockClear();
  });

  it("initialize a instance.", () => {
    // Given
    const options: graph.Options = {} as graph.Options;

    // When
    const actual = ExcelOnlineClient.init(options);

    // Then
    expect(actual).toBeInstanceOf(ExcelOnlineClient);
    expect(graph.Client.init).toHaveBeenCalledWith(options);
    expect(graph.Client.initWithMiddleware).not.toHaveBeenCalled();
  });

  it("initialize a instance with middleware.", () => {
    // Given
    const clientOptions: graph.ClientOptions = {} as graph.ClientOptions;

    // When
    const actual = ExcelOnlineClient.initWithMiddleware(clientOptions);

    // Then
    expect(actual).toBeInstanceOf(ExcelOnlineClient);
    expect(graph.Client.init).not.toHaveBeenCalled();
    expect(graph.Client.initWithMiddleware).toHaveBeenCalledWith(clientOptions);
  });
});
