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

export type StarterTemplateId = "blank" | "delivery-note" | "pallet-label";

export const starterTemplates: Array<{ id: StarterTemplateId; name: string; description: string }> = [
  { id: "blank", name: "Pusty A4", description: "Czysta strona do budowania od zera" },
  { id: "delivery-note", name: "Dokument WZ", description: "Nagłówek, dane odbiorcy, tabela pozycji" },
  { id: "pallet-label", name: "Etykieta paletowa", description: "Format 100 x 150 mm z dużym kodem" },
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
      },
    ],
  };
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
