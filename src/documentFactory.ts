import type { DocumentElement, DocumentTemplate, ElementKind, MarkerColumn } from "./types";

const now = () => new Date().toISOString();

export const defaultMarkerColumns: MarkerColumn[] = [
  { id: "code", label: "#", field: "code" },
  { id: "data", label: "Data", field: "data" },
  { id: "length", label: "Length", field: "length" },
  { id: "format", label: "Format", field: "format" },
  { id: "origin", label: "Origin", field: "origin" },
  { id: "managementRules", label: "Management rules", field: "managementRules" },
];

export type StarterTemplateId = "blank" | "delivery-note" | "pallet-label" | "goods-receipt-advice";

export const starterTemplates: Array<{ id: StarterTemplateId; name: string; description: string }> = [
  { id: "blank", name: "Pusty A4", description: "Czysta strona do budowania od zera" },
  { id: "delivery-note", name: "Dokument WZ", description: "Nagłówek, dane odbiorcy, tabela pozycji" },
  { id: "pallet-label", name: "Etykieta paletowa", description: "Format 100 x 150 mm z dużym kodem" },
  { id: "goods-receipt-advice", name: "Awizacja przyjecia", description: "A4 poziomo z tabela, znacznikami i wymiarami" },
];

export function createTemplate(name = "Nowy dokument"): DocumentTemplate {
  return createDeliveryNoteTemplate(name);
}

export function createTemplateFromStarter(starterId: StarterTemplateId): DocumentTemplate {
  if (starterId === "blank") {
    return createBlankTemplate("Pusty dokument");
  }

  if (starterId === "pallet-label") {
    return createPalletLabelTemplate("Etykieta paletowa");
  }

  if (starterId === "goods-receipt-advice") {
    return createGoodsReceiptAdviceTemplate("Awizacja przyjecia");
  }

  return createDeliveryNoteTemplate("Dokument WZ");
}

function createBlankTemplate(name: string): DocumentTemplate {
  return {
    id: crypto.randomUUID(),
    name,
    updatedAt: now(),
    page: {
      format: "A4",
      widthMm: 210,
      heightMm: 297,
      marginMm: 8,
    },
    markerColumns: defaultMarkerColumns,
    elements: [],
  };
}

function createDeliveryNoteTemplate(name: string): DocumentTemplate {
  return {
    id: crypto.randomUUID(),
    name,
    updatedAt: now(),
    page: {
      format: "A4",
      widthMm: 210,
      heightMm: 297,
      marginMm: 8,
    },
    markerColumns: defaultMarkerColumns,
    elements: [
      {
        id: crypto.randomUUID(),
        kind: "text",
        name: "Nagłówek",
        xMm: 14,
        yMm: 14,
        widthMm: 90,
        heightMm: 10,
        rotation: 0,
        text: "Dokument logistyczny",
        fontSize: 16,
        fontWeight: "700",
        align: "left",
        color: "#172033",
        orientation: "horizontal",
        verticalDirection: "clockwise",
      },
      {
        id: crypto.randomUUID(),
        kind: "box",
        name: "Dane odbiorcy",
        xMm: 14,
        yMm: 34,
        widthMm: 84,
        heightMm: 16,
        rotation: 0,
        borderWidth: 1,
        fill: "transparent",
      },
      {
        id: crypto.randomUUID(),
        kind: "table",
        name: "Tabela pozycji",
        xMm: 14,
        yMm: 56,
        widthMm: 182,
        heightMm: 72,
        rotation: 0,
        rows: 5,
        columns: 4,
        columnWidthsMm: [42, 70, 35, 35],
        rowHeightsMm: [12, 15, 15, 15, 15],
        cells: createTableCells(5, 4, ["SKU", "Nazwa", "Ilość", "Uwagi"]),
        fontSize: 12,
        borderWidth: 1,
      },
      {
        id: crypto.randomUUID(),
        kind: "barcode",
        name: "Kod kreskowy",
        xMm: 132,
        yMm: 16,
        widthMm: 54,
        heightMm: 18,
        rotation: 0,
        value: "LDD-000001",
        symbology: "code128",
      },
    ],
  };
}

function createPalletLabelTemplate(name: string): DocumentTemplate {
  return {
    id: crypto.randomUUID(),
    name,
    updatedAt: now(),
    page: {
      format: "Label100x150",
      widthMm: 100,
      heightMm: 150,
      marginMm: 5,
    },
    markerColumns: defaultMarkerColumns,
    elements: [
      {
        id: crypto.randomUUID(),
        kind: "text",
        name: "Tytuł",
        xMm: 8,
        yMm: 8,
        widthMm: 84,
        heightMm: 14,
        rotation: 0,
        text: "PALETA",
        fontSize: 24,
        fontWeight: "700",
        align: "center",
        color: "#172033",
        orientation: "horizontal",
        verticalDirection: "clockwise",
      },
      {
        id: crypto.randomUUID(),
        kind: "barcode",
        name: "Kod palety",
        xMm: 12,
        yMm: 34,
        widthMm: 76,
        heightMm: 28,
        rotation: 0,
        value: "PAL-000001",
        symbology: "code128",
      },
      {
        id: crypto.randomUUID(),
        kind: "table",
        name: "Dane palety",
        xMm: 8,
        yMm: 76,
        widthMm: 84,
        heightMm: 52,
        rotation: 0,
        rows: 4,
        columns: 2,
        columnWidthsMm: [34, 50],
        rowHeightsMm: [13, 13, 13, 13],
        cells: createTableCells(4, 2, ["Pole", "Wartość"]),
        fontSize: 12,
        borderWidth: 1,
      },
    ],
  };
}

function createGoodsReceiptAdviceTemplate(name: string): DocumentTemplate {
  const tableHeaders = [
    "Nr",
    "SKU",
    "EAN",
    "Oczekiwano\nkart.",
    "Oczekiwano\nszt.",
    "Szt./\nkart.",
    "Nowe\nSKU",
    "Czy jest\nInner",
    "Opakowanie\njednostkowe",
    "Ilosc\nw innerze",
    "Scan\nPallets",
    "Kartony\n/paleta",
    "Typ\nnosnika",
    "Pallet size",
    "Ilosc palet do\nulozenia",
  ];
  const tableRows = [
    ["1", "100977668", "5059340004426", "20", "600", "30", "", "T", "EA", "6", "", "15", "EUR", "", "2"],
    ["2", "100979771", "5059340004242", "10", "300", "30", "", "T", "EA", "5", "", "17", "EUR", "", "1"],
    ["3", "100979775", "5059340003801", "61", "3050", "50", "", "T", "EA", "5", "", "40", "EUR", "", "2"],
    ["4", "100979776", "5059340004433", "11", "550", "50", "", "T", "EA", "5", "", "8", "EUR", "", "2"],
    ["5", "100979778", "5059340004259", "1", "40", "40", "", "T", "EA", "4", "", "20", "EUR", "", "1"],
    ["6", "100979781", "5059340003931", "2", "100", "50", "", "T", "EA", "5", "", "18", "EUR", "", "1"],
    ["7", "100979782", "5059340004198", "20", "500", "25", "", "T", "EA", "5", "", "24", "EUR", "", "1"],
    ["8", "101022806", "5059340004419", "8", "400", "50", "", "T", "EA", "5", "", "40", "EUR", "", "1"],
  ];

  return {
    id: crypto.randomUUID(),
    name,
    updatedAt: now(),
    page: {
      format: "Custom",
      widthMm: 297,
      heightMm: 210,
      marginMm: 8,
    },
    markerColumns: defaultMarkerColumns,
    elements: [
      createArrowElement("Wymiar szerokosci", 20, 10, 260, 8, "horizontal", "29,7 cm", "above"),
      createArrowElement("Wymiar wysokosci", 16, 14, 8, 178, "vertical", "21 cm", "left"),
      createTextElement("Tytul dokumentu", 122, 31, 70, 8, "Awizacja Przyjecia 25 / 000008867", 10, "700", "center"),
      createBarcodeElement("Kod awizacji", 190, 26, 64, 16, "000008867"),
      createTextElement("Numer strony", 263, 27, 20, 9, "Strona   5/   6", 8, "700", "left"),
      ...createLabeledBox("Purchase Order", "0137769920", 42, 44, 115, 17),
      ...createLabeledBox("Nr Inbound Delivery", "8076618602", 42, 64, 115, 17),
      ...createLabeledBox("Dostawca", "TOPWIN HARDWARE & TOOLS MANUFA\nEAST OF NANHE ROAD\n\n515000 GUANGDONG\nCN CHINY", 171, 44, 114, 26),
      createTextElement(
        "Dane dostawy",
        172,
        73,
        124,
        19,
        "Nr kontenera/Nr rejestracyjny        EITU1973521\nData dostawy                         4/08/2025            R105\nGodzina dostawy                      6:00\nWarehouse                            DC2",
        7.4,
        "700",
        "left",
      ),
      createTextElement("Scan palety", 224, 85, 44, 7, "PALLET TO SCAN", 8, "700", "left"),
      createTableElement("Tabela pozycji", 42, 92, 244, 80, tableHeaders, tableRows),
      createTextElement("Suma", 67, 173, 20, 5, "Suma", 7.2, "700", "center"),
      createTextElement("Suma kartonow", 105, 173, 18, 5, "133", 7.2, "700", "center"),
      createTextElement("Suma sztuk", 126, 173, 18, 5, "5540", 7.2, "700", "center"),
      createTextElement("Suma palet", 274, 173, 12, 5, "11", 7.2, "700", "center"),
      createTableElement("Uwagi", 55, 185, 200, 8, ["Uwagi", ""], [], [36, 164]),
      createTextElement("Stopka lewa", 44, 198, 65, 5, "Awizacja Przyjecia 25 / 000008867", 6.4, "700", "left"),
      createTextElement("Stopka prawa", 238, 198, 48, 5, "CPRD09    25/09/2025 13:48:41 R800SWAA", 5.8, "400", "right"),
      ...createGoodsReceiptMarkers(),
    ],
  };
}

function createTextElement(
  name: string,
  xMm: number,
  yMm: number,
  widthMm: number,
  heightMm: number,
  text: string,
  fontSize: number,
  fontWeight: "400" | "700",
  align: "left" | "center" | "right",
): Extract<DocumentElement, { kind: "text" }> {
  return {
    id: crypto.randomUUID(),
    kind: "text",
    name,
    xMm,
    yMm,
    widthMm,
    heightMm,
    rotation: 0,
    text,
    fontSize,
    fontWeight,
    align,
    color: "#172033",
    orientation: "horizontal",
    verticalDirection: "clockwise",
  };
}

function createLabeledBox(
  label: string,
  value: string,
  xMm: number,
  yMm: number,
  widthMm: number,
  heightMm: number,
): DocumentElement[] {
  return [
    {
      id: crypto.randomUUID(),
      kind: "box",
      name: label,
      xMm,
      yMm,
      widthMm,
      heightMm,
      rotation: 0,
      borderWidth: 1.25,
      fill: "transparent",
    },
    createTextElement(`${label} label`, xMm + 2, yMm + 1.5, widthMm - 4, 4, label, 7.4, "700", "left"),
    createTextElement(`${label} value`, xMm + 2, yMm + 6, widthMm - 4, heightMm - 7, value, 7, "400", "left"),
  ];
}

function createBarcodeElement(
  name: string,
  xMm: number,
  yMm: number,
  widthMm: number,
  heightMm: number,
  value: string,
): Extract<DocumentElement, { kind: "barcode" }> {
  return {
    id: crypto.randomUUID(),
    kind: "barcode",
    name,
    xMm,
    yMm,
    widthMm,
    heightMm,
    rotation: 0,
    value,
    symbology: "code128",
  };
}

function createArrowElement(
  name: string,
  xMm: number,
  yMm: number,
  widthMm: number,
  heightMm: number,
  orientation: "horizontal" | "vertical",
  labelText: string,
  labelSide: "above" | "below" | "left" | "right",
): Extract<DocumentElement, { kind: "arrow" }> {
  return {
    id: crypto.randomUUID(),
    kind: "arrow",
    name,
    xMm,
    yMm,
    widthMm,
    heightMm,
    rotation: 0,
    orientation,
    strokeWidth: 1.3,
    headSize: 5,
    head: "both",
    headStyle: "open",
    showLabel: true,
    labelText,
    labelFontSize: 7,
    labelUnit: "cm",
    labelSide,
    labelPosition: 50,
  };
}

function createTableElement(
  name: string,
  xMm: number,
  yMm: number,
  widthMm: number,
  heightMm: number,
  headers: string[],
  rows: string[][],
  columnWidthsMm?: number[],
): Extract<DocumentElement, { kind: "table" }> {
  const columns = headers.length;
  const visibleRows = rows.length > 0 ? rows.length + 1 : 1;
  const rowHeightsMm = rows.length > 0 ? [11, ...Array.from({ length: rows.length }, () => 6.4)] : [heightMm];
  const cells = [
    ...headers.map((header) => ({ text: header, align: "center" as const, verticalAlign: "middle" as const })),
    ...rows.flatMap((row) =>
      Array.from({ length: columns }, (_, index) => ({
        text: row[index] ?? "",
        align: index === 1 || index === 2 ? ("left" as const) : ("center" as const),
        verticalAlign: "middle" as const,
      })),
    ),
  ];

  return {
    id: crypto.randomUUID(),
    kind: "table",
    name,
    xMm,
    yMm,
    widthMm,
    heightMm,
    rotation: 0,
    rows: visibleRows,
    columns,
    columnWidthsMm:
      columnWidthsMm ?? [8, 19, 29, 21, 21, 11, 12, 14, 21, 15, 12, 13, 11, 18, 19],
    rowHeightsMm,
    cells,
    fontSize: rows.length > 0 ? 7.2 : 9,
    borderWidth: 1.4,
  };
}

function createGoodsReceiptMarkers(): Array<Extract<DocumentElement, { kind: "marker" }>> {
  const positions: Array<[string, number, number]> = [
    ["H1", 185, 29],
    ["H2", 67, 46],
    ["H3", 73, 66],
    ["H4", 191, 46],
    ["H5", 167, 72],
    ["H6", 167, 76],
    ["H7", 167, 80],
    ["H8", 167, 84],
    ["H9", 224, 78],
    ["H10", 223, 85],
    ["H11", 263, 23],
    ["H12", 45, 203],
    ["H13", 244, 203],
    ["C1", 46, 148],
    ["C2", 64, 149],
    ["C3", 93, 149],
    ["C4", 114, 149],
    ["C5", 134, 149],
    ["C6", 146, 149],
    ["C7", 159, 149],
    ["C8", 172, 149],
    ["C9", 193, 149],
    ["C10", 208, 149],
    ["C11", 220, 149],
    ["C12", 232, 149],
    ["C13", 243, 149],
    ["C14", 261, 149],
    ["C15", 278, 149],
    ["C16", 67, 178],
    ["C17", 106, 178],
    ["C18", 127, 178],
    ["C19", 73, 189],
    ["C20", 99, 189],
  ];

  return positions.map(([code, xMm, yMm]) => ({
    id: crypto.randomUUID(),
    kind: "marker",
    name: `Znacznik ${code}`,
    xMm,
    yMm,
    widthMm: 10,
    heightMm: 5,
    rotation: 0,
    code,
    data: "",
    length: "",
    format: "",
    origin: "",
    managementRules: "",
    customFields: {},
    fontSize: 7.8,
    fontWeight: "400",
  }));
}

export function createElement(kind: ElementKind): DocumentElement {
  const base = {
    id: crypto.randomUUID(),
    xMm: 22,
    yMm: 32,
    widthMm: 60,
    heightMm: 18,
    rotation: 0,
  };

  if (kind === "text") {
    return {
      ...base,
      kind,
      name: "Tekst",
      text: "Nowy tekst",
      fontSize: 11,
      fontWeight: "400",
      align: "left",
      color: "#172033",
      orientation: "horizontal",
      verticalDirection: "clockwise",
    };
  }

  if (kind === "marker") {
    return {
      ...base,
      kind,
      name: "Znacznik",
      widthMm: 14,
      heightMm: 7,
      code: "H1",
      data: "",
      length: "",
      format: "",
      origin: "",
      managementRules: "",
      customFields: {},
      fontSize: 12,
      fontWeight: "700",
    };
  }

  if (kind === "box") {
    return {
      ...base,
      kind,
      name: "Ramka",
      borderWidth: 1,
      fill: "transparent",
    };
  }

  if (kind === "line") {
    return {
      ...base,
      kind,
      name: "Linia",
      heightMm: 3,
      borderWidth: 1,
      orientation: "horizontal",
    };
  }

  if (kind === "table") {
    return {
      ...base,
      kind,
      name: "Tabela",
      widthMm: 90,
      heightMm: 42,
      rows: 3,
      columns: 3,
      columnWidthsMm: [30, 30, 30],
      rowHeightsMm: [14, 14, 14],
      cells: createTableCells(3, 3, ["Kol. 1", "Kol. 2", "Kol. 3"]),
      fontSize: 12,
      borderWidth: 1,
    };
  }

  if (kind === "image") {
    return {
      ...base,
      kind,
      name: "Logo / obraz",
      widthMm: 42,
      heightMm: 24,
      src: "",
    };
  }

  if (kind === "arrow") {
    return {
      ...base,
      kind,
      name: "Strzałka",
      widthMm: 60,
      heightMm: 14,
      orientation: "horizontal",
      strokeWidth: 1.5,
      headSize: 5,
      head: "both",
      headStyle: "triangle",
      showLabel: true,
      labelText: "",
      labelFontSize: 7,
      labelUnit: "mm",
      labelSide: "above",
      labelPosition: 50,
    };
  }

  return {
    ...base,
    kind,
    name: "Kod",
    widthMm: 54,
    heightMm: 18,
    value: "LDD-000001",
    symbology: "code128",
  };
}

function createTableCells(rows: number, columns: number, headers: string[] = []) {
  return Array.from({ length: rows * columns }, (_, index) => ({
    text: index < columns ? headers[index] ?? "" : "",
    align: index < columns ? ("center" as const) : ("left" as const),
    verticalAlign: "middle" as const,
  }));
}
