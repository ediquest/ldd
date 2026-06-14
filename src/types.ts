export type PageFormat = "A4" | "A5" | "Letter" | "Label100x150" | "Custom";

export type MeasurementUnit = "mm" | "cm" | "in";

export type DocumentPage = {
  format: PageFormat;
  widthMm: number;
  heightMm: number;
  marginMm: number;
  background?: PageBackground;
};

export type PageBackground = {
  src: string;
  xMm: number;
  yMm: number;
  widthMm: number;
  heightMm: number;
  opacity: number;
};

export type ElementKind = "text" | "box" | "line" | "table" | "image" | "barcode" | "arrow" | "marker";

type ElementBase = {
  id: string;
  kind: ElementKind;
  xMm: number;
  yMm: number;
  widthMm: number;
  heightMm: number;
  rotation: number;
  name: string;
  locked?: boolean;
};

export type TextElement = ElementBase & {
  kind: "text";
  text: string;
  fontSize: number;
  fontWeight: "400" | "700";
  align: "left" | "center" | "right";
  color?: string;
  orientation?: "horizontal" | "vertical";
  verticalDirection?: "clockwise" | "counterclockwise";
};

export type MarkerElement = ElementBase & {
  kind: "marker";
  code: string;
  data: string;
  length: string;
  format: string;
  origin: string;
  managementRules: string;
  customFields?: Record<string, string>;
  fontSize: number;
  fontWeight: "400" | "700";
};

export type MarkerColumn = {
  id: string;
  label: string;
  field: "code" | "data" | "length" | "format" | "origin" | "managementRules" | "custom";
  customFieldId?: string;
};

export type BoxElement = ElementBase & {
  kind: "box";
  borderWidth: number;
  fill: string;
};

export type LineElement = ElementBase & {
  kind: "line";
  borderWidth: number;
  orientation?: "horizontal" | "vertical";
};

export type TableElement = ElementBase & {
  kind: "table";
  rows: number;
  columns: number;
  columnWidthsMm?: number[];
  rowHeightsMm?: number[];
  cells?: TableCell[];
  fontSize: number;
  borderWidth?: number;
};

export type TableCell = {
  text: string;
  align: "left" | "center" | "right";
  verticalAlign: "top" | "middle" | "bottom";
};

export type ImageElement = ElementBase & {
  kind: "image";
  src: string;
};

export type BarcodeElement = ElementBase & {
  kind: "barcode";
  value: string;
  symbology:
    | "code128"
    | "gs1-128"
    | "code39"
    | "ean13"
    | "itf14"
    | "upca"
    | "qr"
    | "datamatrix"
    | "pdf417";
};

export type ArrowElement = ElementBase & {
  kind: "arrow";
  orientation: "horizontal" | "vertical";
  strokeWidth: number;
  headSize?: number;
  head: "end" | "start" | "both";
  headStyle: "triangle" | "bar" | "open";
  showLabel: boolean;
  labelText?: string;
  labelFontSize?: number;
  labelUnit: "mm" | "cm" | "in";
  labelSide: "above" | "below" | "left" | "right";
  labelPosition: number;
};

export type DocumentElement =
  | TextElement
  | MarkerElement
  | BoxElement
  | LineElement
  | TableElement
  | ImageElement
  | BarcodeElement
  | ArrowElement;

export type DocumentTemplate = {
  id: string;
  name: string;
  updatedAt: string;
  page: DocumentPage;
  markerColumns?: MarkerColumn[];
  elements: DocumentElement[];
};
