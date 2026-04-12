import { describe, it, expect } from "vitest";
import {
  createUntypedArray,
  createUntypedString,
  createUntypedNumber,
  createUntypedBoolean,
  createUntypedNull,
} from "@microsoft/kiota-abstractions";
import { type WorkbookRange } from "@microsoft/msgraph-sdk/models/index.js";
import { convertRangeToCells } from "./range-converter";
import { Cell } from "./cell";

function makeRange(overrides: Partial<WorkbookRange> = {}): WorkbookRange {
  return { values: null, valueTypes: null, text: null, numberFormat: null, ...overrides };
}

describe("convertRangeToCells", () => {
  it("valuesがnullのとき空配列を返す", () => {
    const result = convertRangeToCells(makeRange());
    expect(result).toEqual([]);
  });

  it("valuesが配列でないとき空配列を返す", () => {
    const result = convertRangeToCells(makeRange({ values: createUntypedString("bad") }));
    expect(result).toEqual([]);
  });

  it("文字列・数値・真偽値・nullを含む範囲を正しく変換する", () => {
    const range = makeRange({
      values: createUntypedArray([
        createUntypedArray([
          createUntypedString("Alice"),
          createUntypedNumber(30),
          createUntypedBoolean(true),
          createUntypedNull(),
        ]),
      ]),
      valueTypes: createUntypedArray([
        createUntypedArray([
          createUntypedString("String"),
          createUntypedString("Integer"),
          createUntypedString("Boolean"),
          createUntypedString("Empty"),
        ]),
      ]),
      text: createUntypedArray([
        createUntypedArray([
          createUntypedString("Alice"),
          createUntypedString("30"),
          createUntypedString("TRUE"),
          createUntypedString(""),
        ]),
      ]),
      numberFormat: createUntypedArray([
        createUntypedArray([
          createUntypedString("@"),
          createUntypedString("0"),
          createUntypedString("General"),
          createUntypedString("General"),
        ]),
      ]),
    });

    const result = convertRangeToCells(range);

    expect(result).toHaveLength(1);
    expect(result[0]).toHaveLength(4);

    expect(result[0][0]).toBeInstanceOf(Cell);
    expect(result[0][0].value).toBe("Alice");
    expect(result[0][0].valueType).toBe("String");
    expect(result[0][0].text).toBe("Alice");
    expect(result[0][0].numberFormat).toBe("@");

    expect(result[0][1].value).toBe(30);
    expect(result[0][1].valueType).toBe("Integer");

    expect(result[0][2].value).toBe(true);
    expect(result[0][2].valueType).toBe("Boolean");

    expect(result[0][3].value).toBeNull();
    expect(result[0][3].valueType).toBe("Empty");
  });

  it("複数行を正しく変換する", () => {
    const range = makeRange({
      values: createUntypedArray([
        createUntypedArray([createUntypedString("A"), createUntypedNumber(1)]),
        createUntypedArray([createUntypedString("B"), createUntypedNumber(2)]),
      ]),
    });

    const result = convertRangeToCells(range);

    expect(result).toHaveLength(2);
    expect(result[0][0].value).toBe("A");
    expect(result[0][1].value).toBe(1);
    expect(result[1][0].value).toBe("B");
    expect(result[1][1].value).toBe(2);
  });

  it("valueTypes・text・numberFormatがnullのときフォールバック値を使う", () => {
    const range = makeRange({
      values: createUntypedArray([
        createUntypedArray([createUntypedString("X")]),
      ]),
    });

    const result = convertRangeToCells(range);

    expect(result[0][0].value).toBe("X");
    expect(result[0][0].valueType).toBe("Unknown");
    expect(result[0][0].text).toBeNull();
    expect(result[0][0].numberFormat).toBeNull();
  });

  it("未知のvalueType文字列はUnknownになる", () => {
    const range = makeRange({
      values: createUntypedArray([
        createUntypedArray([createUntypedString("X")]),
      ]),
      valueTypes: createUntypedArray([
        createUntypedArray([createUntypedString("InvalidType")]),
      ]),
    });

    const result = convertRangeToCells(range);

    expect(result[0][0].valueType).toBe("Unknown");
  });
});
