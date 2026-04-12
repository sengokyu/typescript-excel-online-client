export type ValueType =
  | "Unknown"
  | "Empty"
  | "String"
  | "Integer"
  | "Double"
  | "Boolean"
  | "Error";

/**
 * Represents a cell in an Excel worksheet.
 */
export class Cell {
  constructor(
    private datum: {
      value: string | number | boolean | null;
      valueType: ValueType;
      text: unknown;
      numberFormat: string | null;
    },
  ) {}

  get valueType(): ValueType {
    return this.datum.valueType;
  }

  get value(): string | number | boolean | null {
    return this.datum.value;
  }

  get text(): unknown {
    return this.datum.text;
  }

  get numberFormat(): string | null {
    return this.datum.numberFormat;
  }
}
