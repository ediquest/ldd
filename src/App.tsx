import {
  CSSProperties,
  ChangeEvent,
  FocusEvent,
  KeyboardEvent as ReactKeyboardEvent,
  MouseEvent,
  PointerEvent,
  RefObject,
  ReactNode,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import {
  AlignmentType,
  Document as DocxDocument,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import {
  Box,
  ArrowRight,
  ChevronDown,
  ChevronUp,
  FileText,
  Image,
  LayoutTemplate,
  Lock,
  Copy,
  Redo2,
  Minus,
  Pencil,
  Plus,
  QrCode,
  Ruler,
  Save,
  Tag,
  Table2,
  Type,
  Trash2,
  Unlock,
  Undo2,
} from "lucide-react";
import {
  createElement,
  createTemplate,
  createTemplateFromStarter,
  defaultMarkerColumns,
  starterTemplates,
} from "./documentFactory";
import { clearDraft, deleteTemplate, loadDraft, loadDrafts, loadTemplates, saveDraft, saveTemplate } from "./db";
import type {
  DocumentElement,
  DocumentPage,
  DocumentTemplate,
  ElementKind,
  MeasurementUnit,
  PageFormat,
  TableElement,
  MarkerColumn,
} from "./types";

const PX_PER_MM = 3.7795275591;
const MIN_ZOOM = 0.55;
const MAX_ZOOM = 3;
const ZOOM_STEP = 0.05;

const units: Record<MeasurementUnit, { label: string; suffix: string; factorMm: number; step: number }> = {
  mm: { label: "Milimetry", suffix: "mm", factorMm: 1, step: 1 },
  cm: { label: "Centymetry", suffix: "cm", factorMm: 10, step: 0.1 },
  in: { label: "Cale", suffix: "in", factorMm: 25.4, step: 0.01 },
};

const pagePresets: Array<{ format: PageFormat; label: string; widthMm: number; heightMm: number; marginMm: number }> = [
  { format: "A4", label: "A4 210 x 297 mm", widthMm: 210, heightMm: 297, marginMm: 8 },
  { format: "A5", label: "A5 148 x 210 mm", widthMm: 148, heightMm: 210, marginMm: 7 },
  { format: "Letter", label: "Letter 8.5 x 11 in", widthMm: 215.9, heightMm: 279.4, marginMm: 8 },
  { format: "Label100x150", label: "Etykieta 100 x 150 mm", widthMm: 100, heightMm: 150, marginMm: 5 },
];

const barcodeSymbologies: Array<{
  value: Extract<DocumentElement, { kind: "barcode" }>["symbology"];
  label: string;
  group: "1D" | "2D";
}> = [
  { value: "code128", label: "Code 128", group: "1D" },
  { value: "gs1-128", label: "GS1-128 / UCC/EAN-128", group: "1D" },
  { value: "code39", label: "Code 39", group: "1D" },
  { value: "ean13", label: "EAN-13", group: "1D" },
  { value: "itf14", label: "ITF-14", group: "1D" },
  { value: "upca", label: "UPC-A", group: "1D" },
  { value: "qr", label: "QR Code", group: "2D" },
  { value: "datamatrix", label: "Data Matrix", group: "2D" },
  { value: "pdf417", label: "PDF417", group: "2D" },
];

type DragState = {
  id: string;
  startX: number;
  startY: number;
  originalX: number;
  originalY: number;
  duplicateOnDrag: boolean;
  duplicatedId?: string;
  historyCaptured?: boolean;
};

type ResizeState = {
  id: string;
  startX: number;
  startY: number;
  originalX: number;
  originalY: number;
  originalWidth: number;
  originalHeight: number;
  originalFontSize?: number;
  historyCaptured?: boolean;
};

type TableSelection = {
  tableId: string;
  axis: "column" | "row";
  index: number;
};

type TableCellSelection = {
  tableId: string;
  row: number;
  column: number;
};

type TableResizeState = {
  tableId: string;
  axis: "column" | "row";
  index: number;
  startX: number;
  startY: number;
  originalSizes: number[];
  historyCaptured?: boolean;
};

type CanvasPanState = {
  pointerId: number;
  startX: number;
  startY: number;
  scrollLeft: number;
  scrollTop: number;
  container: HTMLDivElement;
};

type ZoomAnchorState = {
  clientX: number;
  clientY: number;
  pageX: number;
  pageY: number;
  container: HTMLDivElement;
};

type AssistSettings = {
  snapToElements: boolean;
  showGuides: boolean;
  stepMove: boolean;
  stepMm: number;
};

type ActiveGuides = {
  vertical: number[];
  horizontal: number[];
};

type MarkerColumnKey = string;

type MarkerColumnWidths = Partial<Record<MarkerColumnKey, number>>;

type ContextMenuState = {
  screenX: number;
  screenY: number;
  pageX: number;
  pageY: number;
} | null;

const defaultMarkerColumnWidths: MarkerColumnWidths = {
  code: 80,
  data: 180,
  length: 110,
  format: 140,
  origin: 180,
  managementRules: 320,
};
const HISTORY_LIMIT = 30;

type DocumentUpdateOptions = {
  history?: "push" | "skip";
};

type PageOrientation = "portrait" | "landscape";

export default function App() {
  const [templates, setTemplates] = useState<DocumentTemplate[]>([]);
  const [activeTemplate, setActiveTemplate] = useState<DocumentTemplate>(() => createTemplate());
  const [undoStack, setUndoStack] = useState<DocumentTemplate[]>([]);
  const [redoStack, setRedoStack] = useState<DocumentTemplate[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedTablePart, setSelectedTablePart] = useState<TableSelection | null>(null);
  const [selectedTableCell, setSelectedTableCell] = useState<TableCellSelection | null>(null);
  const [editingTableCell, setEditingTableCell] = useState<TableCellSelection | null>(null);
  const [status, setStatus] = useState("Gotowe do pracy");
  const [autosaveReady, setAutosaveReady] = useState(false);
  const [zoom, setZoom] = useState(0.86);
  const [unit, setUnit] = useState<MeasurementUnit>("mm");
  const [assistSettings, setAssistSettings] = useState<AssistSettings>({
    snapToElements: true,
    showGuides: true,
    stepMove: false,
    stepMm: 1,
  });
  const [activeGuides, setActiveGuides] = useState<ActiveGuides>({ vertical: [], horizontal: [] });
  const [contextMenu, setContextMenu] = useState<ContextMenuState>(null);
  const [markerPanelOpen, setMarkerPanelOpen] = useState(true);
  const [starterTemplatesOpen, setStarterTemplatesOpen] = useState(false);
  const [savedTemplatesOpen, setSavedTemplatesOpen] = useState(true);
  const [markerColumnWidths, setMarkerColumnWidths] = useState<MarkerColumnWidths>(() => loadMarkerColumnWidths());
  const [isSpacePanMode, setIsSpacePanMode] = useState(false);
  const [isCanvasPanning, setIsCanvasPanning] = useState(false);
  const isSpacePanModeRef = useRef(false);
  const zoomRef = useRef(zoom);
  const selectedIdRef = useRef<string | null>(selectedId);
  const dragRef = useRef<DragState | null>(null);
  const resizeRef = useRef<ResizeState | null>(null);
  const tableResizeRef = useRef<TableResizeState | null>(null);
  const canvasPanRef = useRef<CanvasPanState | null>(null);
  const zoomAnchorRef = useRef<ZoomAnchorState | null>(null);
  const canvasScrollRef = useRef<HTMLDivElement | null>(null);
  const markerPanelRef = useRef<HTMLElement | null>(null);
  const pageRef = useRef<HTMLDivElement | null>(null);
  const activeTemplateRef = useRef(activeTemplate);
  const skipNextAutosaveRef = useRef(true);
  const autosaveTimerRef = useRef<number | null>(null);

  const selectedElement = useMemo(
    () => activeTemplate.elements.find((element) => element.id === selectedId) ?? null,
    [activeTemplate.elements, selectedId],
  );
  const markers = useMemo(
    () => activeTemplate.elements.filter((element): element is Extract<DocumentElement, { kind: "marker" }> => element.kind === "marker"),
    [activeTemplate.elements],
  );
  const markerColumns = useMemo(
    () => activeTemplate.markerColumns?.length ? activeTemplate.markerColumns : defaultMarkerColumns,
    [activeTemplate.markerColumns],
  );
  const allElementsLocked = activeTemplate.elements.length > 0 && activeTemplate.elements.every((element) => element.locked);

  useEffect(() => {
    activeTemplateRef.current = activeTemplate;
  }, [activeTemplate]);

  useEffect(() => {
    zoomRef.current = zoom;
  }, [zoom]);

  useEffect(() => {
    selectedIdRef.current = selectedId;
  }, [selectedId]);

  useLayoutEffect(() => {
    const anchor = zoomAnchorRef.current;
    const page = pageRef.current;
    if (!anchor || !page) {
      return;
    }

    const pageRect = page.getBoundingClientRect();
    const anchoredScreenX = pageRect.left + anchor.pageX * zoom;
    const anchoredScreenY = pageRect.top + anchor.pageY * zoom;
    anchor.container.scrollLeft += anchoredScreenX - anchor.clientX;
    anchor.container.scrollTop += anchoredScreenY - anchor.clientY;

    const canvasPan = canvasPanRef.current;
    if (canvasPan?.container === anchor.container) {
      canvasPan.startX = anchor.clientX;
      canvasPan.startY = anchor.clientY;
      canvasPan.scrollLeft = anchor.container.scrollLeft;
      canvasPan.scrollTop = anchor.container.scrollTop;
    }

    zoomAnchorRef.current = null;
  }, [zoom]);

  function clearAutosaveTimer() {
    if (!autosaveTimerRef.current) {
      return;
    }

    window.clearTimeout(autosaveTimerRef.current);
    autosaveTimerRef.current = null;
  }

  function saveDraftNow(template: DocumentTemplate = activeTemplateRef.current) {
    clearAutosaveTimer();
    return saveDraft(template)
      .then(() => setStatus(`Zapisano automatycznie: ${new Date().toLocaleTimeString("pl-PL")}`))
      .catch(() => setStatus("Nie udało się zapisać szkicu"));
  }

  function resetHistory() {
    setUndoStack([]);
    setRedoStack([]);
  }

  function pushUndoSnapshot(template: DocumentTemplate = activeTemplateRef.current) {
    setUndoStack((current) => [...current.slice(-HISTORY_LIMIT + 1), structuredClone(template)]);
    setRedoStack([]);
  }

  function setSpacePanMode(active: boolean) {
    isSpacePanModeRef.current = active;
    setIsSpacePanMode(active);
  }

  function applyTemplateState(template: DocumentTemplate) {
    activeTemplateRef.current = template;
    setActiveTemplate(template);
    setSelectedId((current) => (template.elements.some((element) => element.id === current) ? current : null));
    setSelectedTablePart((current) => (current && template.elements.some((element) => element.id === current.tableId) ? current : null));
    setSelectedTableCell((current) => (current && template.elements.some((element) => element.id === current.tableId) ? current : null));
    setEditingTableCell(null);
    setActiveGuides({ vertical: [], horizontal: [] });
  }

  function undoDocumentChange() {
    setUndoStack((current) => {
      const previousTemplate = current[current.length - 1];
      if (!previousTemplate) {
        return current;
      }

      setRedoStack((redoCurrent) => [...redoCurrent.slice(-HISTORY_LIMIT + 1), structuredClone(activeTemplateRef.current)]);
      applyTemplateState(structuredClone(previousTemplate));
      setStatus("Cofnięto zmianę");
      return current.slice(0, -1);
    });
  }

  function redoDocumentChange() {
    setRedoStack((current) => {
      const nextTemplate = current[current.length - 1];
      if (!nextTemplate) {
        return current;
      }

      setUndoStack((undoCurrent) => [...undoCurrent.slice(-HISTORY_LIMIT + 1), structuredClone(activeTemplateRef.current)]);
      applyTemplateState(structuredClone(nextTemplate));
      setStatus("Ponowiono zmianę");
      return current.slice(0, -1);
    });
  }

  useEffect(() => {
    Promise.all([loadTemplates(), loadDrafts()])
      .then(([storedTemplates, drafts]) => {
        setTemplates(storedTemplates);
        const newestDraft = drafts
          .filter((draft) => {
            const savedTemplate = storedTemplates.find((template) => template.id === draft.template.id);
            return !savedTemplate || draft.template.updatedAt.localeCompare(savedTemplate.updatedAt) > 0;
          })
          .sort((a, b) => b.savedAt.localeCompare(a.savedAt))[0];
        const shouldRestoreDraft =
          newestDraft &&
          window.confirm(
            `Znaleziono automatycznie zapisany szkic dokumentu "${newestDraft.template.name}". Przywrócić niezapisane zmiany?`,
          );
        const templateToOpen = shouldRestoreDraft ? newestDraft.template : storedTemplates[0];

        if (newestDraft && !shouldRestoreDraft) {
          void clearDraft(newestDraft.id);
        }

        if (templateToOpen) {
          setActiveTemplate(templateToOpen);
          setSelectedId(templateToOpen.elements[0]?.id ?? null);
          resetHistory();
          setStatus(shouldRestoreDraft ? "Przywrócono szkic autosave" : "Gotowe do pracy");
        }
        skipNextAutosaveRef.current = true;
        setAutosaveReady(true);
      })
      .catch(() => {
        setStatus("Nie udało się odczytać IndexedDB");
        setAutosaveReady(true);
      });
  }, []);

  useEffect(() => {
    if (!autosaveReady) {
      return;
    }

    if (skipNextAutosaveRef.current) {
      skipNextAutosaveRef.current = false;
      return;
    }

    clearAutosaveTimer();

    setStatus("Zapisywanie automatyczne...");
    autosaveTimerRef.current = window.setTimeout(() => {
      autosaveTimerRef.current = null;
      void saveDraftNow(activeTemplate);
    }, 900);

    return clearAutosaveTimer;
  }, [activeTemplate, autosaveReady]);

  useEffect(() => {
    function onBeforeUnload(event: BeforeUnloadEvent) {
      if (!autosaveTimerRef.current) {
        return;
      }

      event.preventDefault();
      event.returnValue = "";
    }

    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, []);

  useEffect(() => {
    const styleId = "ldd-print-page-size";
    let styleElement = document.getElementById(styleId) as HTMLStyleElement | null;
    const pageWidth = `${activeTemplate.page.widthMm}mm`;
    const pageHeight = `${activeTemplate.page.heightMm}mm`;

    if (!styleElement) {
      styleElement = document.createElement("style");
      styleElement.id = styleId;
      document.head.appendChild(styleElement);
    }

    document.documentElement.style.setProperty("--page-width", pageWidth);
    document.documentElement.style.setProperty("--page-height", pageHeight);

    styleElement.textContent = `
      @page {
        size: ${pageWidth} ${pageHeight};
        margin: 0;
      }
    `;
  }, [activeTemplate.page.heightMm, activeTemplate.page.widthMm]);

  useEffect(() => {
    localStorage.setItem("ldd-marker-column-widths", JSON.stringify(markerColumnWidths));
  }, [markerColumnWidths]);

  useEffect(() => {
    const container = canvasScrollRef.current;
    if (!container) {
      return;
    }
    const scrollContainer = container;

    function onWheel(event: WheelEvent) {
      if (!isSpacePanModeRef.current) {
        adjustSelectedElementWithWheel(event);
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      zoomCanvasAtPointer(scrollContainer, event.clientX, event.clientY, event.deltaY);
    }

    scrollContainer.addEventListener("wheel", onWheel, { passive: false });
    return () => scrollContainer.removeEventListener("wheel", onWheel);
  }, []);

  useEffect(() => {
    function isEditableTarget() {
      const activeElement = document.activeElement;
      return (
        activeElement instanceof HTMLInputElement ||
        activeElement instanceof HTMLTextAreaElement ||
        activeElement instanceof HTMLSelectElement ||
        activeElement instanceof HTMLElement && activeElement.isContentEditable
      );
    }

    function onKeyDown(event: KeyboardEvent) {
      const isEditing = isEditableTarget();

      if (event.code === "Space" && !isEditing && !event.ctrlKey && !event.metaKey && !event.altKey) {
        event.preventDefault();
        setSpacePanMode(true);
        return;
      }

      if (event.key === "Delete" && selectedId && !isEditing) {
        event.preventDefault();
        deleteElement(selectedId);
      }

      if ((event.ctrlKey || event.metaKey) && !isEditing) {
        const key = event.key.toLowerCase();
        if (key === "z" && !event.shiftKey) {
          event.preventDefault();
          undoDocumentChange();
        }

        if (key === "y" || (key === "z" && event.shiftKey)) {
          event.preventDefault();
          redoDocumentChange();
        }
      }

      if (!isEditing && selectedId && ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(event.key)) {
        event.preventDefault();
        const directionByKey: Record<string, "up" | "down" | "left" | "right"> = {
          ArrowUp: "up",
          ArrowDown: "down",
          ArrowLeft: "left",
          ArrowRight: "right",
        };
        moveSelectedElementByKeyboard(directionByKey[event.key], event.shiftKey);
      }

      if (event.key === "Escape") {
        setContextMenu(null);
      }
    }

    function onKeyUp(event: KeyboardEvent) {
      if (event.code === "Space") {
        setSpacePanMode(false);
        stopCanvasPan();
      }
    }

    function onWindowBlur() {
      setSpacePanMode(false);
      stopCanvasPan();
    }

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    window.addEventListener("blur", onWindowBlur);
    return () => {
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("keyup", onKeyUp);
      window.removeEventListener("blur", onWindowBlur);
    };
  }, [assistSettings.stepMm, assistSettings.stepMove, selectedId]);

  useEffect(() => {
    function closeContextMenu() {
      setContextMenu(null);
    }

    window.addEventListener("click", closeContextMenu);
    window.addEventListener("resize", closeContextMenu);
    return () => {
      window.removeEventListener("click", closeContextMenu);
      window.removeEventListener("resize", closeContextMenu);
    };
  }, []);

  function updateDocument(updater: (current: DocumentTemplate) => DocumentTemplate, options: DocumentUpdateOptions = {}) {
    setActiveTemplate((current) => {
      if (options.history !== "skip") {
        setUndoStack((history) => [...history.slice(-HISTORY_LIMIT + 1), structuredClone(current)]);
        setRedoStack([]);
      }
      const nextTemplate = updater({ ...current, updatedAt: new Date().toISOString() });
      activeTemplateRef.current = nextTemplate;
      return nextTemplate;
    });
  }

  function updateElement(id: string, patch: Partial<DocumentElement>, options: DocumentUpdateOptions = {}) {
    updateDocument((current) => ({
      ...current,
      elements: current.elements.map((element) =>
        element.id === id ? ({ ...element, ...patch } as DocumentElement) : element,
      ),
    }), options);
  }

  function updateTableLayout(id: string, patch: Partial<TableElement>, options: DocumentUpdateOptions = {}) {
    updateDocument((current) => ({
      ...current,
      elements: current.elements.map((element) => {
        if (element.id !== id || element.kind !== "table") {
          return element;
        }

        return normalizeTable({
          ...element,
          ...patch,
        });
      }),
    }), options);
  }

  function updateTableCell(tableId: string, row: number, column: number, patch: Partial<NonNullable<TableElement["cells"]>[number]>) {
    updateDocument((current) => ({
      ...current,
      elements: current.elements.map((element) => {
        if (element.id !== tableId || element.kind !== "table") {
          return element;
        }

        const table = normalizeTable(element);
        const index = row * table.columns + column;
        const cells = [...(table.cells ?? [])];
        cells[index] = {
          ...createEmptyCell(),
          ...cells[index],
          ...patch,
        };

        return {
          ...table,
          cells,
        };
      }),
    }));
  }

  function moveSelectedElementByKeyboard(direction: "up" | "down" | "left" | "right", largeStep: boolean) {
    const element = activeTemplateRef.current.elements.find((item) => item.id === selectedId);
    if (!element || element.locked) {
      return;
    }

    const baseStep = assistSettings.stepMove ? assistSettings.stepMm : 1;
    const stepMm = largeStep ? baseStep * 10 : baseStep;
    const deltaX = direction === "left" ? -stepMm : direction === "right" ? stepMm : 0;
    const deltaY = direction === "up" ? -stepMm : direction === "down" ? stepMm : 0;

    updateElement(element.id, {
      xMm: Math.max(0, roundMm(element.xMm + deltaX)),
      yMm: Math.max(0, roundMm(element.yMm + deltaY)),
    });
  }

  function duplicateElement(element: DocumentElement, options: DocumentUpdateOptions = {}) {
    const duplicatedElement = {
      ...element,
      id: crypto.randomUUID(),
      name: `${element.name} copy`,
      xMm: element.xMm + 4,
      yMm: element.yMm + 4,
    } as DocumentElement;

    if (duplicatedElement.kind === "marker" && element.kind === "marker") {
      duplicatedElement.code = nextMarkerCode(element.code, markers.map((item) => item.code));
      duplicatedElement.name = "Znacznik";
    }

    updateDocument((current) => ({
      ...current,
      elements: [...current.elements, duplicatedElement],
    }), options);
    setSelectedId(duplicatedElement.id);
    if (duplicatedElement.kind === "marker") {
      setMarkerPanelOpen(true);
    }

    return duplicatedElement.id;
  }

  function deleteElement(id: string) {
    updateDocument((current) => ({ ...current, elements: current.elements.filter((element) => element.id !== id) }));
    setSelectedId(null);
    setSelectedTablePart(null);
    setSelectedTableCell(null);
    setEditingTableCell(null);
    setStatus("Usunięto element");
  }

  function toggleElementLock(id: string) {
    const element = activeTemplate.elements.find((item) => item.id === id);
    updateElement(id, { locked: !element?.locked });
  }

  function toggleAllElementLocks() {
    if (activeTemplate.elements.length === 0) {
      return;
    }

    const locked = !allElementsLocked;
    updateDocument((current) => ({
      ...current,
      elements: current.elements.map((element) => ({ ...element, locked })),
    }));
    setSelectedId(null);
    setSelectedTablePart(null);
    setSelectedTableCell(null);
    setEditingTableCell(null);
    setStatus(locked ? "Zablokowano wszystkie elementy" : "Odblokowano wszystkie elementy");
  }

  function updatePage(patch: Partial<DocumentPage>) {
    updateDocument((current) => ({
      ...current,
      page: {
        ...current.page,
        ...patch,
      },
    }));
  }

  function updateMarkerColumns(columns: MarkerColumn[]) {
    updateDocument((current) => ({
      ...current,
      markerColumns: columns,
    }));
  }

  function renameMarkerColumn(columnId: string, label: string) {
    updateMarkerColumns(markerColumns.map((column) => (column.id === columnId ? { ...column, label } : column)));
  }

  function addMarkerColumn() {
    const id = `custom-${crypto.randomUUID()}`;
    updateMarkerColumns([
      ...markerColumns,
      {
        id,
        label: "Nowa kolumna",
        field: "custom",
        customFieldId: id,
      },
    ]);
    setMarkerColumnWidths((current) => ({ ...current, [id]: 160 }));
  }

  function deleteMarkerColumn(columnId: string) {
    updateMarkerColumns(markerColumns.filter((column) => column.id !== columnId));
    setMarkerColumnWidths((current) => {
      const nextWidths = { ...current };
      delete nextWidths[columnId];
      return nextWidths;
    });
  }

  function applyPagePreset(format: PageFormat) {
    if (format === "Custom") {
      updatePage({ format });
      return;
    }

    const preset = pagePresets.find((item) => item.format === format);
    if (!preset) {
      return;
    }
    const orientedSize = orientPageSize(preset, getPageOrientation(activeTemplate.page));

    updatePage({
      format: preset.format,
      widthMm: orientedSize.widthMm,
      heightMm: orientedSize.heightMm,
      marginMm: preset.marginMm,
    });
  }

  function changePageOrientation(orientation: PageOrientation) {
    const currentOrientation = getPageOrientation(activeTemplate.page);
    if (orientation === currentOrientation) {
      return;
    }

    const preset = pagePresets.find((item) => item.format === activeTemplate.page.format);
    if (preset) {
      const orientedSize = orientPageSize(preset, orientation);
      updatePage({
        widthMm: orientedSize.widthMm,
        heightMm: orientedSize.heightMm,
      });
      return;
    }

    updatePage({
      widthMm: activeTemplate.page.heightMm,
      heightMm: activeTemplate.page.widthMm,
    });
  }

  function addElement(kind: ElementKind, position?: { xMm: number; yMm: number }) {
    const element = createElement(kind);
    const placedElement = position
      ? ({
          ...element,
          xMm: roundMm(position.xMm),
          yMm: roundMm(position.yMm),
        } as DocumentElement)
      : element;

    updateDocument((current) => ({ ...current, elements: [...current.elements, placedElement] }));
    setSelectedId(placedElement.id);
    setSelectedTablePart(null);
    setSelectedTableCell(null);
    setEditingTableCell(null);
    setContextMenu(null);
    if (placedElement.kind === "marker") {
      setMarkerPanelOpen(true);
    }
  }

  async function persistTemplate() {
    clearAutosaveTimer();
    await saveTemplate(activeTemplate);
    await clearDraft(activeTemplate.id);
    const storedTemplates = await loadTemplates();
    setTemplates(storedTemplates);
    setStatus(`Zapisano: ${new Date().toLocaleTimeString("pl-PL")}`);
  }

  async function openSavedTemplate(template: DocumentTemplate) {
    await saveDraftNow();
    const draft = await loadDraft(template.id);
    const templateToOpen =
      draft && draft.template.updatedAt.localeCompare(template.updatedAt) > 0 ? draft.template : template;

    skipNextAutosaveRef.current = true;
    setActiveTemplate(templateToOpen);
    activeTemplateRef.current = templateToOpen;
    setSelectedId(templateToOpen.elements[0]?.id ?? null);
    setSelectedTablePart(null);
    setSelectedTableCell(null);
    setEditingTableCell(null);
    resetHistory();
    setStatus(draft && templateToOpen === draft.template ? "Wczytano szkic autosave dokumentu" : "Wczytano zapisany dokument");
  }

  async function duplicateTemplate(template: DocumentTemplate) {
    const copiedTemplate: DocumentTemplate = {
      ...structuredClone(template),
      id: crypto.randomUUID(),
      name: `${template.name} kopia`,
      updatedAt: new Date().toISOString(),
      elements: template.elements.map((element) => ({
        ...structuredClone(element),
        id: crypto.randomUUID(),
      })) as DocumentElement[],
    };

    await saveTemplate(copiedTemplate);
    const storedTemplates = await loadTemplates();
    setTemplates(storedTemplates);
    setActiveTemplate(copiedTemplate);
    activeTemplateRef.current = copiedTemplate;
    setSelectedId(copiedTemplate.elements[0]?.id ?? null);
    resetHistory();
    setStatus("Skopiowano dokument");
  }

  async function removeSavedTemplate(template: DocumentTemplate) {
    const confirmed = window.confirm(`Usunąć zapisany dokument "${template.name}"?`);
    if (!confirmed) {
      return;
    }

    await deleteTemplate(template.id);
    await clearDraft(template.id);
    const storedTemplates = await loadTemplates();
    setTemplates(storedTemplates);

    if (activeTemplate.id === template.id) {
      const nextTemplate = storedTemplates[0] ?? createTemplate();
      setActiveTemplate(nextTemplate);
      activeTemplateRef.current = nextTemplate;
      setSelectedId(nextTemplate.elements[0]?.id ?? null);
      resetHistory();
    }

    setStatus("Usunięto zapisany dokument");
  }

  function printDocument() {
    document.documentElement.style.setProperty("--page-width", `${activeTemplate.page.widthMm}mm`);
    document.documentElement.style.setProperty("--page-height", `${activeTemplate.page.heightMm}mm`);
    document.body.classList.add("print-fit");
    const cleanup = () => {
      document.body.classList.remove("print-fit");
      window.removeEventListener("afterprint", cleanup);
    };

    window.addEventListener("afterprint", cleanup);
    requestAnimationFrame(() => window.print());
  }

  async function exportPdf() {
    if (!pageRef.current) {
      return;
    }

    document.body.classList.add("export-mode");
    setStatus("Eksportuję PDF...");

    try {
      const canvas = await html2canvas(pageRef.current, {
        backgroundColor: "#ffffff",
        scale: 2,
      });
      const pdf = new jsPDF({
        orientation: activeTemplate.page.widthMm > activeTemplate.page.heightMm ? "landscape" : "portrait",
        unit: "mm",
        format: [activeTemplate.page.widthMm, activeTemplate.page.heightMm],
      });
      const imageData = canvas.toDataURL("image/png");

      pdf.addImage(imageData, "PNG", 0, 0, activeTemplate.page.widthMm, activeTemplate.page.heightMm);
      addMarkerPagesToPdf(pdf, markers, markerColumns, activeTemplate.page.widthMm, activeTemplate.page.heightMm);
      pdf.save(`${sanitizeFileName(activeTemplate.name)}.pdf`);
      setStatus("PDF gotowy");
    } finally {
      document.body.classList.remove("export-mode");
    }
  }

  async function exportDocx() {
    if (!pageRef.current) {
      return;
    }

    document.body.classList.add("export-mode");
    setStatus("Eksportuję DOCX...");

    try {
      const canvas = await html2canvas(pageRef.current, {
        backgroundColor: "#ffffff",
        scale: 2,
      });
      const imageBlob = await new Promise<Blob>((resolve, reject) => {
        canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error("Nie udało się utworzyć obrazu dokumentu"))), "image/png");
      });
      const imageData = await imageBlob.arrayBuffer();
      const doc = createDocxExport(imageData, activeTemplate.name, markers, markerColumns);
      const blob = await Packer.toBlob(doc);
      downloadBlob(blob, `${sanitizeFileName(activeTemplate.name)}.docx`);
      setStatus("DOCX gotowy");
    } finally {
      document.body.classList.remove("export-mode");
    }
  }

  async function copyMarkersToClipboard() {
    const html = markerTableToHtml(markers, markerColumns);
    const text = markerTableToPlainText(markers, markerColumns);

    if (navigator.clipboard && "ClipboardItem" in window) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ]);
    } else {
      await navigator.clipboard.writeText(text);
    }

    setStatus("Skopiowano tabelę znaczników");
  }

  async function createNewTemplate() {
    await saveDraftNow();
    const freshTemplate = createTemplate(`Dokument ${templates.length + 1}`);
    skipNextAutosaveRef.current = true;
    setActiveTemplate(freshTemplate);
    activeTemplateRef.current = freshTemplate;
    setSelectedId(freshTemplate.elements[0]?.id ?? null);
    resetHistory();
    setStatus("Utworzono nowy dokument");
  }

  async function createFromStarter(starterId: Parameters<typeof createTemplateFromStarter>[0]) {
    await saveDraftNow();
    const freshTemplate = createTemplateFromStarter(starterId);
    skipNextAutosaveRef.current = true;
    setActiveTemplate(freshTemplate);
    activeTemplateRef.current = freshTemplate;
    setSelectedId(freshTemplate.elements[0]?.id ?? null);
    resetHistory();
    setStatus("Wczytano szablon startowy");
  }

  function openContextMenu(event: MouseEvent<HTMLDivElement>) {
    event.preventDefault();
    const rect = event.currentTarget.getBoundingClientRect();
    const pageX = (event.clientX - rect.left) / PX_PER_MM / zoom;
    const pageY = (event.clientY - rect.top) / PX_PER_MM / zoom;

    setContextMenu({
      screenX: event.clientX,
      screenY: event.clientY,
      pageX: Math.max(0, Math.min(activeTemplate.page.widthMm, pageX)),
      pageY: Math.max(0, Math.min(activeTemplate.page.heightMm, pageY)),
    });
  }

  function startCanvasPan(event: PointerEvent<HTMLElement>, container: HTMLDivElement) {
    event.preventDefault();
    setContextMenu(null);
    stopInteractions();
    container.setPointerCapture(event.pointerId);
    canvasPanRef.current = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      scrollLeft: container.scrollLeft,
      scrollTop: container.scrollTop,
      container,
    };
    setIsCanvasPanning(true);
  }

  function zoomCanvasAtPointer(container: HTMLDivElement, clientX: number, clientY: number, deltaY: number) {
    const page = pageRef.current;
    if (!page) {
      return;
    }

    const pageRect = page.getBoundingClientRect();
    const currentZoom = zoomRef.current;
    const pageX = (clientX - pageRect.left) / currentZoom;
    const pageY = (clientY - pageRect.top) / currentZoom;
    const direction = deltaY > 0 ? -1 : 1;
    const nextZoom = clampZoom(currentZoom + direction * ZOOM_STEP);

    if (nextZoom === currentZoom) {
      return;
    }

    zoomAnchorRef.current = {
      clientX,
      clientY,
      pageX,
      pageY,
      container,
    };
    zoomRef.current = nextZoom;
    setZoom(nextZoom);
  }

  function adjustSelectedElementWithWheel(event: WheelEvent) {
    const selectedElement = activeTemplateRef.current.elements.find((element) => element.id === selectedIdRef.current);
    if (!selectedElement) {
      return;
    }

    const direction = event.deltaY > 0 ? -1 : 1;

    if (selectedElement.kind === "text" || selectedElement.kind === "marker") {
      event.preventDefault();
      event.stopPropagation();
      updateElement(selectedElement.id, {
        fontSize: Math.max(4, Number((selectedElement.fontSize + direction).toFixed(1))),
      });
      return;
    }

    if (selectedElement.kind === "box" || selectedElement.kind === "table") {
      event.preventDefault();
      event.stopPropagation();
      updateElement(selectedElement.id, {
        borderWidth: Math.max(0, Number(((selectedElement.borderWidth ?? 1) + direction * 0.1).toFixed(1))),
      } as Partial<DocumentElement>);
    }
  }

  function onPointerMove(event: PointerEvent<HTMLDivElement>) {
    const canvasPan = canvasPanRef.current;
    const tableResize = tableResizeRef.current;
    const drag = dragRef.current;
    const resize = resizeRef.current;

    if (canvasPan) {
      canvasPan.container.scrollLeft = canvasPan.scrollLeft - (event.clientX - canvasPan.startX);
      canvasPan.container.scrollTop = canvasPan.scrollTop - (event.clientY - canvasPan.startY);
      return;
    }

    if (tableResize) {
      const table = activeTemplate.elements.find(
        (element): element is TableElement => element.id === tableResize.tableId && element.kind === "table",
      );
      if (!table) {
        return;
      }

      const delta =
        tableResize.axis === "column"
          ? (event.clientX - tableResize.startX) / PX_PER_MM / zoom
          : (event.clientY - tableResize.startY) / PX_PER_MM / zoom;
      if (!tableResize.historyCaptured && Math.abs(delta) > 0.1) {
        pushUndoSnapshot();
        tableResize.historyCaptured = true;
      }
      const minSize = tableResize.axis === "column" ? 5 : 4;
      const sizes = resizeTableSizes(tableResize.originalSizes, tableResize.index, delta, minSize);

      if (tableResize.axis === "column") {
        updateTableLayout(table.id, { columnWidthsMm: sizes }, { history: "skip" });
      } else {
        updateTableLayout(table.id, { rowHeightsMm: sizes }, { history: "skip" });
      }
      return;
    }

    if (drag) {
      const deltaX = (event.clientX - drag.startX) / PX_PER_MM / zoom;
      const deltaY = (event.clientY - drag.startY) / PX_PER_MM / zoom;
      const sourceElement = activeTemplate.elements.find((element) => element.id === drag.id);

      if (!drag.historyCaptured && (Math.abs(deltaX) > 0.1 || Math.abs(deltaY) > 0.1)) {
        pushUndoSnapshot();
        drag.historyCaptured = true;
      }

      if (drag.historyCaptured && drag.duplicateOnDrag && !drag.duplicatedId && sourceElement) {
        drag.duplicatedId = duplicateElement(sourceElement, { history: "skip" });
      }

      const movingId = drag.duplicatedId ?? drag.id;
      const movingElement = activeTemplate.elements.find((element) => element.id === movingId) ?? sourceElement;

      if (!movingElement) {
        return;
      }

      const nextPosition = resolveAssistedPosition({
        element: movingElement,
        elements: activeTemplate.elements,
        xMm: drag.originalX + deltaX,
        yMm: drag.originalY + deltaY,
        settings: assistSettings,
      });

      updateElement(
        movingId,
        {
          xMm: nextPosition.xMm,
          yMm: nextPosition.yMm,
        },
        { history: "skip" },
      );
      setActiveGuides(assistSettings.showGuides ? nextPosition.guides : { vertical: [], horizontal: [] });
    }

    if (resize) {
      const element = activeTemplate.elements.find((item) => item.id === resize.id);
      const deltaX = (event.clientX - resize.startX) / PX_PER_MM / zoom;
      const deltaY = (event.clientY - resize.startY) / PX_PER_MM / zoom;
      if (!resize.historyCaptured && (Math.abs(deltaX) > 0.1 || Math.abs(deltaY) > 0.1)) {
        pushUndoSnapshot();
        resize.historyCaptured = true;
      }
      const nextWidth = Math.max(5, roundMm(resize.originalWidth + deltaX));
      const nextHeight =
        element?.kind === "line" || element?.kind === "arrow"
          ? element.orientation === "vertical"
            ? Math.max(5, roundMm(resize.originalHeight + deltaY))
            : resize.originalHeight
          : Math.max(1, roundMm(resize.originalHeight + deltaY));
      const assistedSize =
        element
          ? resolveAssistedSize({
              element,
              elements: activeTemplate.elements,
              widthMm: nextWidth,
              heightMm: nextHeight,
              settings: assistSettings,
            })
          : { widthMm: nextWidth, heightMm: nextHeight, guides: { vertical: [], horizontal: [] } };
      const finalHeight =
        element?.kind === "line" || element?.kind === "arrow"
          ? element.orientation === "vertical"
            ? assistedSize.heightMm
            : resize.originalHeight
          : assistedSize.heightMm;
      const fontPatch =
        (element?.kind === "text" || element?.kind === "marker") && resize.originalFontSize
          ? {
              fontSize: Math.max(6, Math.round(resize.originalFontSize * (finalHeight / Math.max(1, resize.originalHeight)))),
            }
          : {};

      updateElement(
        resize.id,
        {
          widthMm:
            (element?.kind === "line" || element?.kind === "arrow") && element.orientation === "vertical"
              ? resize.originalWidth
              : assistedSize.widthMm,
          heightMm:
            element?.kind === "line" || element?.kind === "arrow"
              ? element.orientation === "vertical"
                ? finalHeight
                : resize.originalHeight
              : finalHeight,
          ...fontPatch,
        },
        { history: "skip" },
      );
      setActiveGuides(assistSettings.showGuides ? assistedSize.guides : { vertical: [], horizontal: [] });
    }
  }

  function stopInteractions() {
    dragRef.current = null;
    resizeRef.current = null;
    tableResizeRef.current = null;
    stopCanvasPan();
    setActiveGuides({ vertical: [], horizontal: [] });
  }

  function stopCanvasPan() {
    const canvasPan = canvasPanRef.current;
    if (canvasPan?.container.hasPointerCapture(canvasPan.pointerId)) {
      canvasPan.container.releasePointerCapture(canvasPan.pointerId);
    }
    canvasPanRef.current = null;
    setIsCanvasPanning(false);
  }

  function uploadImage(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file || !selectedElement || selectedElement.kind !== "image") {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      updateElement(selectedElement.id, { src: String(reader.result) });
    };
    reader.readAsDataURL(file);
  }

  function uploadPageBackground(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      updatePage({
        background: {
          src: String(reader.result),
          xMm: 0,
          yMm: 0,
          widthMm: activeTemplate.page.widthMm,
          heightMm: activeTemplate.page.heightMm,
          opacity: 0.35,
        },
      });
    };
    reader.readAsDataURL(file);
  }

  function focusCanvasScroll() {
    canvasScrollRef.current?.focus({ preventScroll: true });
  }

  return (
    <main
      className={markerPanelOpen ? "app-shell marker-panel-visible" : "app-shell"}
      style={
        {
          "--page-width": `${activeTemplate.page.widthMm}mm`,
          "--page-height": `${activeTemplate.page.heightMm}mm`,
        } as CSSProperties
      }
    >
      <aside className="left-panel">
        <div className="brand">
          <FileText size={22} />
          <div>
            <strong>LDD</strong>
            <span>Designer dokumentów</span>
          </div>
        </div>

        <section className="panel-section">
          <button className="primary-button" onClick={createNewTemplate}>
            <Plus size={16} /> Nowy dokument
          </button>
          <button className="secondary-button" onClick={() => void persistTemplate()}>
            <Save size={16} /> Zapisz lokalnie
          </button>
        </section>

        <section className="panel-section starter-list">
          <button
            className="accordion-trigger"
            type="button"
            onClick={() => setStarterTemplatesOpen((current) => !current)}
          >
            <span>Szablony startowe</span>
            <strong>{starterTemplatesOpen ? "−" : "+"}</strong>
          </button>
          <div className={starterTemplatesOpen ? "accordion-content open" : "accordion-content"}>
            {starterTemplates.map((template) => (
              <button key={template.id} className="starter-item" onClick={() => createFromStarter(template.id)}>
                <LayoutTemplate size={16} />
                <span>
                  <strong>{template.name}</strong>
                  <small>{template.description}</small>
                </span>
              </button>
            ))}
          </div>
        </section>

        <section className="panel-section">
          <h2>Narzędzia pomocnicze</h2>
          <AssistSettingsPanel settings={assistSettings} onChange={setAssistSettings} />
        </section>

        <section className="panel-section">
          <h2>Elementy</h2>
          <div className="tool-grid">
            <ToolButton icon={<Type size={16} />} label="Tekst" onClick={() => addElement("text")} />
            <ToolButton icon={<Box size={16} />} label="Ramka" onClick={() => addElement("box")} />
            <ToolButton icon={<Minus size={16} />} label="Linia" onClick={() => addElement("line")} />
            <ToolButton icon={<Table2 size={16} />} label="Tabela" onClick={() => addElement("table")} />
            <ToolButton icon={<Image size={16} />} label="Obraz" onClick={() => addElement("image")} />
            <ToolButton icon={<QrCode size={16} />} label="Kod" onClick={() => addElement("barcode")} />
            <ToolButton icon={<ArrowRight size={16} />} label="Strzałka" onClick={() => addElement("arrow")} />
            <ToolButton icon={<Tag size={16} />} label="Znacznik" onClick={() => addElement("marker")} />
          </div>
        </section>

        <section className="panel-section saved-list">
          <button
            className="accordion-trigger"
            type="button"
            onClick={() => setSavedTemplatesOpen((current) => !current)}
          >
            <span>Zapisane</span>
            <strong>{savedTemplatesOpen ? "−" : "+"}</strong>
          </button>
          <div className={savedTemplatesOpen ? "accordion-content saved-content open" : "accordion-content saved-content"}>
            {templates.length === 0 ? (
              <p>Brak zapisanych dokumentów.</p>
            ) : (
              templates.map((template) => (
                <div key={template.id} className={template.id === activeTemplate.id ? "template-item active" : "template-item"}>
                  <button
                    className="template-open-button"
                    onClick={() => void openSavedTemplate(template)}
                  >
                    <strong>{template.name}</strong>
                    <span>{new Date(template.updatedAt).toLocaleString("pl-PL")}</span>
                  </button>
                  <button
                    className="template-copy-button"
                    title="Kopiuj dokument"
                    aria-label={`Kopiuj dokument ${template.name}`}
                    onClick={(event) => {
                      event.stopPropagation();
                      void duplicateTemplate(template);
                    }}
                  >
                    <Copy size={15} />
                  </button>
                  <button
                    className="template-delete-button"
                    title="Usuń zapisany dokument"
                    aria-label={`Usuń zapisany dokument ${template.name}`}
                    onClick={(event) => {
                      event.stopPropagation();
                      void removeSavedTemplate(template);
                    }}
                  >
                    <Trash2 size={15} />
                  </button>
                </div>
              ))
            )}
          </div>
        </section>
      </aside>

      <section className="workspace">
        <header className="top-bar">
          <input
            aria-label="Nazwa dokumentu"
            value={activeTemplate.name}
            onChange={(event) => updateDocument((current) => ({ ...current, name: event.target.value }))}
          />
          <div className="top-actions">
            <span>{status}</span>
            <button
              type="button"
              title="Cofnij (Ctrl+Z)"
              aria-label="Cofnij"
              disabled={undoStack.length === 0}
              onClick={undoDocumentChange}
            >
              <Undo2 size={15} />
            </button>
            <button
              type="button"
              title="Ponów (Ctrl+Y)"
              aria-label="Ponów"
              disabled={redoStack.length === 0}
              onClick={redoDocumentChange}
            >
              <Redo2 size={15} />
            </button>
            <button
              type="button"
              title={allElementsLocked ? "Odblokuj wszystkie elementy" : "Zablokuj wszystkie elementy"}
              aria-label={allElementsLocked ? "Odblokuj wszystkie elementy" : "Zablokuj wszystkie elementy"}
              disabled={activeTemplate.elements.length === 0}
              onClick={toggleAllElementLocks}
            >
              {allElementsLocked ? <Unlock size={15} /> : <Lock size={15} />}
            </button>
            <label>
              Zoom
              <input
                type="range"
                min={MIN_ZOOM}
                max={MAX_ZOOM}
                step={ZOOM_STEP}
                value={zoom}
                onChange={(event) => {
                  const nextZoom = Number(event.target.value);
                  zoomRef.current = nextZoom;
                  setZoom(nextZoom);
                }}
                onKeyDown={(event) => {
                  if (event.key === " ") {
                    event.preventDefault();
                    event.currentTarget.blur();
                    focusCanvasScroll();
                    setSpacePanMode(true);
                  }
                }}
                onKeyUp={(event) => {
                  if (event.key === "Enter" || event.key === " ") {
                    event.currentTarget.blur();
                    focusCanvasScroll();
                  }
                }}
                onPointerUp={(event) => {
                  event.currentTarget.blur();
                  focusCanvasScroll();
                }}
              />
            </label>
            <button onClick={printDocument}>Drukuj</button>
            <button onClick={() => void exportPdf()}>PDF</button>
            <button onClick={() => void exportDocx()}>DOCX</button>
          </div>
        </header>

        <div
          ref={canvasScrollRef}
          tabIndex={-1}
          className={`canvas-scroll ${isSpacePanMode ? "pan-mode" : ""} ${isCanvasPanning ? "panning" : ""}`}
          onPointerMove={onPointerMove}
          onPointerUp={stopInteractions}
          onPointerLeave={stopInteractions}
          onPointerDown={(event) => {
            if (isSpacePanMode) {
              startCanvasPan(event, event.currentTarget);
              return;
            }

            if (event.target !== event.currentTarget) {
              return;
            }

            setSelectedId(null);
            setSelectedTablePart(null);
            setSelectedTableCell(null);
            setEditingTableCell(null);
          }}
        >
          <div
            ref={pageRef}
            className="page"
            style={{
              width: activeTemplate.page.widthMm * PX_PER_MM * zoom,
              height: activeTemplate.page.heightMm * PX_PER_MM * zoom,
              "--page-width": `${activeTemplate.page.widthMm}mm`,
              "--page-height": `${activeTemplate.page.heightMm}mm`,
            } as CSSProperties}
            onContextMenu={openContextMenu}
            onPointerDown={() => {
              if (isSpacePanMode) {
                return;
              }

              setSelectedId(null);
              setSelectedTablePart(null);
              setSelectedTableCell(null);
              setEditingTableCell(null);
            }}
          >
            <div
              className="safe-margin"
              style={{
                inset: activeTemplate.page.marginMm * PX_PER_MM * zoom,
              }}
            />
            {activeTemplate.page.background && (
              <img
                className="page-background-image"
                src={activeTemplate.page.background.src}
                alt=""
                draggable={false}
                style={{
                  left: activeTemplate.page.background.xMm * PX_PER_MM * zoom,
                  top: activeTemplate.page.background.yMm * PX_PER_MM * zoom,
                  width: activeTemplate.page.background.widthMm * PX_PER_MM * zoom,
                  height: activeTemplate.page.background.heightMm * PX_PER_MM * zoom,
                  opacity: activeTemplate.page.background.opacity,
                }}
              />
            )}
            {activeGuides.vertical.map((xMm) => (
              <div
                key={`vertical-${xMm}`}
                className="guide-line vertical"
                style={{ left: xMm * PX_PER_MM * zoom }}
              />
            ))}
            {activeGuides.horizontal.map((yMm) => (
              <div
                key={`horizontal-${yMm}`}
                className="guide-line horizontal"
                style={{ top: yMm * PX_PER_MM * zoom }}
              />
            ))}
            {activeTemplate.elements.map((element) => (
              <DocumentElementView
                key={element.id}
                element={element}
                zoom={zoom}
                selected={element.id === selectedId}
                selectedTablePart={element.id === selectedId ? selectedTablePart : null}
                selectedTableCell={element.id === selectedId ? selectedTableCell : null}
                editingTableCell={element.id === selectedId ? editingTableCell : null}
                onSelect={() => {
                  setSelectedId(element.id);
                  if (element.kind !== "table") {
                    setSelectedTablePart(null);
                    setSelectedTableCell(null);
                    setEditingTableCell(null);
                  }
                }}
                onDragStart={(event) => {
                  event.stopPropagation();
                  if (isSpacePanMode) {
                    const container = canvasScrollRef.current;
                    if (container) {
                      startCanvasPan(event, container);
                    }
                    return;
                  }

                  setSelectedId(element.id);
                  if (element.kind !== "table") {
                    setSelectedTablePart(null);
                    setSelectedTableCell(null);
                    setEditingTableCell(null);
                  }
                  if (element.kind === "marker") {
                    setMarkerPanelOpen(true);
                  }
                  if (element.locked) {
                    return;
                  }
                  dragRef.current = {
                    id: element.id,
                    startX: event.clientX,
                    startY: event.clientY,
                    originalX: element.xMm,
                    originalY: element.yMm,
                    duplicateOnDrag: event.altKey,
                  };
                }}
                onResizeStart={(event) => {
                  event.stopPropagation();
                  if (isSpacePanMode) {
                    const container = canvasScrollRef.current;
                    if (container) {
                      startCanvasPan(event, container);
                    }
                    return;
                  }

                  setSelectedId(element.id);
                  if (element.locked) {
                    return;
                  }
                  resizeRef.current = {
                    id: element.id,
                    startX: event.clientX,
                    startY: event.clientY,
                    originalX: element.xMm,
                    originalY: element.yMm,
                    originalWidth: element.widthMm,
                    originalHeight: element.heightMm,
                    originalFontSize:
                      element.kind === "text" || element.kind === "marker" ? element.fontSize : undefined,
                  };
                }}
                onToggleLock={() => toggleElementLock(element.id)}
                onTableDividerPointerDown={(event, table, axis, index) => {
                  event.stopPropagation();
                  if (isSpacePanMode) {
                    const container = canvasScrollRef.current;
                    if (container) {
                      startCanvasPan(event, container);
                    }
                    return;
                  }

                  setSelectedId(table.id);
                  if (table.locked) {
                    return;
                  }
                  setSelectedTablePart({ tableId: table.id, axis, index });
                  setSelectedTableCell(null);
                  const normalizedTable = normalizeTable(table);
                  tableResizeRef.current = {
                    tableId: table.id,
                    axis,
                    index,
                    startX: event.clientX,
                    startY: event.clientY,
                    originalSizes:
                      axis === "column" ? normalizedTable.columnWidthsMm ?? [] : normalizedTable.rowHeightsMm ?? [],
                  };
                }}
                onTableCellDoubleClick={(event, table, row, column) => {
                  event.stopPropagation();
                  setSelectedId(table.id);
                  setSelectedTablePart(null);
                  setSelectedTableCell({ tableId: table.id, row, column });
                  setEditingTableCell({ tableId: table.id, row, column });
                }}
                onTableCellChange={(table, row, column, text) => updateTableCell(table.id, row, column, { text })}
                onTableCellEditEnd={() => setEditingTableCell(null)}
              />
            ))}
          </div>
        </div>
      </section>

      {contextMenu && (
        <ContextMenu
          menu={contextMenu}
          assistSettings={assistSettings}
          onAssistSettingsChange={setAssistSettings}
          onAddElement={(kind) => addElement(kind, { xMm: contextMenu.pageX, yMm: contextMenu.pageY })}
          onClose={() => setContextMenu(null)}
        />
      )}

      <aside className="right-panel">
        <section className="properties-section">
          <h2>Dokument</h2>
          <PageSettingsPanel
            page={activeTemplate.page}
            unit={unit}
            onUnitChange={setUnit}
            onPageChange={updatePage}
            onPresetChange={applyPagePreset}
            onOrientationChange={changePageOrientation}
            onBackgroundUpload={(event) => uploadPageBackground(event)}
          />
        </section>

        <section className="properties-section">
          <h2>Właściwości</h2>
        {!selectedElement ? (
          <p className="muted">Zaznacz element na stronie.</p>
        ) : (
          <>
            <PropertiesPanel
              element={selectedElement}
              unit={unit}
              selectedTablePart={selectedTablePart}
              selectedTableCell={selectedTableCell}
              onChange={(patch) => updateElement(selectedElement.id, patch)}
              onImageUpload={uploadImage}
            />
            <button className="danger-button" onClick={() => deleteElement(selectedElement.id)}>
              <Trash2 size={16} /> Usuń element
            </button>
          </>
        )}
        </section>
      </aside>

      <MarkerPanel
        panelRef={markerPanelRef}
        markers={markers}
        columns={markerColumns}
        columnWidths={markerColumnWidths}
        selectedMarkerId={selectedElement?.kind === "marker" ? selectedElement.id : null}
        open={markerPanelOpen}
        onOpenChange={setMarkerPanelOpen}
        onSelectMarker={(id) => {
          setMarkerPanelOpen(true);
          setSelectedId(id);
        }}
        onChangeMarker={(id, patch) => updateElement(id, patch)}
        onCommitMarkerEdit={() => void saveDraftNow()}
        onRenameColumn={renameMarkerColumn}
        onAddColumn={addMarkerColumn}
        onDeleteColumn={deleteMarkerColumn}
        onCopyMarkers={() => void copyMarkersToClipboard()}
        onColumnWidthChange={(column, width) =>
          setMarkerColumnWidths((current) => ({ ...current, [column]: Math.max(60, width) }))
        }
      />
    </main>
  );
}

function ToolButton({ icon, label, onClick }: { icon: ReactNode; label: string; onClick: () => void }) {
  return (
    <button className="tool-button" onClick={onClick}>
      {icon}
      {label}
    </button>
  );
}

function ContextMenu({
  menu,
  assistSettings,
  onAssistSettingsChange,
  onAddElement,
  onClose,
}: {
  menu: NonNullable<ContextMenuState>;
  assistSettings: AssistSettings;
  onAssistSettingsChange: (settings: AssistSettings) => void;
  onAddElement: (kind: ElementKind) => void;
  onClose: () => void;
}) {
  const elementTools: Array<{ kind: ElementKind; label: string; icon: ReactNode }> = [
    { kind: "text", label: "Tekst", icon: <Type size={16} /> },
    { kind: "box", label: "Ramka", icon: <Box size={16} /> },
    { kind: "line", label: "Linia", icon: <Minus size={16} /> },
    { kind: "table", label: "Tabela", icon: <Table2 size={16} /> },
    { kind: "image", label: "Obraz", icon: <Image size={16} /> },
    { kind: "barcode", label: "Kod", icon: <QrCode size={16} /> },
    { kind: "arrow", label: "Strzałka", icon: <ArrowRight size={16} /> },
    { kind: "marker", label: "Znacznik", icon: <Tag size={16} /> },
  ];

  return (
    <div
      className="context-menu"
      style={{ left: menu.screenX, top: menu.screenY }}
      onClick={(event) => event.stopPropagation()}
      onContextMenu={(event) => event.preventDefault()}
    >
      <div className="context-menu-header">
        <strong>Dodaj element</strong>
        <span>
          {formatNumber(menu.pageX)} x {formatNumber(menu.pageY)} mm
        </span>
      </div>
      <div className="context-menu-grid">
        {elementTools.map((tool) => (
          <button
            key={tool.kind}
            className="context-menu-item"
            onClick={() => {
              onAddElement(tool.kind);
              onClose();
            }}
          >
            {tool.icon}
            {tool.label}
          </button>
        ))}
      </div>
      <div className="context-menu-section">
        <strong>Narzędzia pomocnicze</strong>
        <AssistSettingsPanel settings={assistSettings} onChange={onAssistSettingsChange} compact />
      </div>
    </div>
  );
}

function AssistSettingsPanel({
  settings,
  onChange,
  compact = false,
}: {
  settings: AssistSettings;
  onChange: (settings: AssistSettings) => void;
  compact?: boolean;
}) {
  function patchSettings(patch: Partial<AssistSettings>) {
    onChange({ ...settings, ...patch });
  }

  return (
    <div className={compact ? "assist-settings compact" : "assist-settings"}>
      <label className="checkbox-row">
        <input
          type="checkbox"
          checked={settings.snapToElements}
          onChange={(event) => patchSettings({ snapToElements: event.target.checked })}
        />
        Przyciąganie do elementów
      </label>
      <label className="checkbox-row">
        <input
          type="checkbox"
          checked={settings.showGuides}
          onChange={(event) => patchSettings({ showGuides: event.target.checked })}
        />
        Linie pomocnicze
      </label>
      <label className="checkbox-row">
        <input
          type="checkbox"
          checked={settings.stepMove}
          onChange={(event) => patchSettings({ stepMove: event.target.checked })}
        />
        Ruch skokowy
      </label>
      <label className="step-field">
        Skok
        <select
          value={settings.stepMm}
          disabled={!settings.stepMove}
          onChange={(event) => patchSettings({ stepMm: Number(event.target.value) })}
        >
          <option value={0.5}>0.5 mm</option>
          <option value={1}>1 mm</option>
          <option value={2}>2 mm</option>
          <option value={5}>5 mm</option>
          <option value={10}>10 mm</option>
        </select>
      </label>
    </div>
  );
}

function PageSettingsPanel({
  page,
  unit,
  onUnitChange,
  onPageChange,
  onPresetChange,
  onOrientationChange,
  onBackgroundUpload,
}: {
  page: DocumentPage;
  unit: MeasurementUnit;
  onUnitChange: (unit: MeasurementUnit) => void;
  onPageChange: (patch: Partial<DocumentPage>) => void;
  onPresetChange: (format: PageFormat) => void;
  onOrientationChange: (orientation: PageOrientation) => void;
  onBackgroundUpload: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
  const [backgroundOpen, setBackgroundOpen] = useState(false);
  const orientation = getPageOrientation(page);

  return (
    <div className="properties-grid">
      <label className="wide">
        Format
        <select value={page.format} onChange={(event) => onPresetChange(event.target.value as PageFormat)}>
          {pagePresets.map((preset) => (
            <option key={preset.format} value={preset.format}>
              {preset.label}
            </option>
          ))}
          <option value="Custom">Własny</option>
        </select>
      </label>
      <div className="orientation-toggle wide">
        <span>Orientacja</span>
        <div>
          <button
            type="button"
            className={orientation === "portrait" ? "active" : ""}
            onClick={() => onOrientationChange("portrait")}
          >
            Pionowo
          </button>
          <button
            type="button"
            className={orientation === "landscape" ? "active" : ""}
            onClick={() => onOrientationChange("landscape")}
          >
            Poziomo
          </button>
        </div>
      </div>
      <label className="wide">
        Jednostka
        <select value={unit} onChange={(event) => onUnitChange(event.target.value as MeasurementUnit)}>
          {Object.entries(units).map(([value, config]) => (
            <option key={value} value={value}>
              {config.label}
            </option>
          ))}
        </select>
      </label>
      <NumberField
        label={`Szer. ${units[unit].suffix}`}
        valueMm={page.widthMm}
        unit={unit}
        minMm={20}
        onChange={(widthMm) => onPageChange({ widthMm, format: "Custom" })}
      />
      <NumberField
        label={`Wys. ${units[unit].suffix}`}
        valueMm={page.heightMm}
        unit={unit}
        minMm={20}
        onChange={(heightMm) => onPageChange({ heightMm, format: "Custom" })}
      />
      <NumberField
        label={`Margines ${units[unit].suffix}`}
        valueMm={page.marginMm}
        unit={unit}
        minMm={0}
        onChange={(marginMm) => onPageChange({ marginMm })}
      />
      <div className="page-summary">
        <Ruler size={15} />
        {formatMeasurement(page.widthMm, unit)} x {formatMeasurement(page.heightMm, unit)}
      </div>
      <div className="page-background-accordion wide">
        <button className="accordion-trigger" type="button" onClick={() => setBackgroundOpen((current) => !current)}>
          <span>Tło dokumentu</span>
          <strong>{backgroundOpen ? "−" : "+"}</strong>
        </button>
        <div className={backgroundOpen ? "page-background-settings open" : "page-background-settings"}>
          <label>
            Obraz tła
            <input type="file" accept="image/*" onChange={onBackgroundUpload} />
          </label>
          {page.background && (
            <>
              <NumberField
                label={`X ${units[unit].suffix}`}
                valueMm={page.background.xMm}
                unit={unit}
                onChange={(xMm) => onPageChange({ background: { ...page.background!, xMm } })}
              />
              <NumberField
                label={`Y ${units[unit].suffix}`}
                valueMm={page.background.yMm}
                unit={unit}
                onChange={(yMm) => onPageChange({ background: { ...page.background!, yMm } })}
              />
              <NumberField
                label={`Szer. ${units[unit].suffix}`}
                valueMm={page.background.widthMm}
                unit={unit}
                minMm={1}
                onChange={(widthMm) => onPageChange({ background: { ...page.background!, widthMm } })}
              />
              <NumberField
                label={`Wys. ${units[unit].suffix}`}
                valueMm={page.background.heightMm}
                unit={unit}
                minMm={1}
                onChange={(heightMm) => onPageChange({ background: { ...page.background!, heightMm } })}
              />
              <label>
                Opacity
                <input
                  type="range"
                  min="0"
                  max="1"
                  step="0.05"
                  value={page.background.opacity}
                  onChange={(event) =>
                    onPageChange({ background: { ...page.background!, opacity: Number(event.target.value) } })
                  }
                />
              </label>
              <button type="button" className="secondary-button" onClick={() => onPageChange({ background: undefined })}>
                Usuń tło
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function DocumentElementView({
  element,
  zoom,
  selected,
  selectedTablePart,
  selectedTableCell,
  editingTableCell,
  onSelect,
  onDragStart,
  onResizeStart,
  onToggleLock,
  onTableDividerPointerDown,
  onTableCellDoubleClick,
  onTableCellChange,
  onTableCellEditEnd,
}: {
  element: DocumentElement;
  zoom: number;
  selected: boolean;
  selectedTablePart: TableSelection | null;
  selectedTableCell: TableCellSelection | null;
  editingTableCell: TableCellSelection | null;
  onSelect: () => void;
  onDragStart: (event: PointerEvent<HTMLDivElement>) => void;
  onResizeStart: (event: PointerEvent<HTMLElement>) => void;
  onToggleLock: () => void;
  onTableDividerPointerDown: (
    event: PointerEvent<HTMLButtonElement>,
    table: TableElement,
    axis: "column" | "row",
    index: number,
  ) => void;
  onTableCellDoubleClick: (
    event: MouseEvent<HTMLSpanElement>,
    table: TableElement,
    row: number,
    column: number,
  ) => void;
  onTableCellChange: (table: TableElement, row: number, column: number, text: string) => void;
  onTableCellEditEnd: () => void;
}) {
  const style = {
    left: element.xMm * PX_PER_MM * zoom,
    top: element.yMm * PX_PER_MM * zoom,
    width: element.widthMm * PX_PER_MM * zoom,
    height: Math.max(1, element.heightMm * PX_PER_MM * zoom),
    transform: `rotate(${element.rotation}deg)`,
  };

  return (
    <div
      className={selected ? "doc-element selected" : "doc-element"}
      style={style}
      onPointerDown={onDragStart}
      onClick={(event) => {
        event.stopPropagation();
        onSelect();
      }}
    >
      <ElementContent
        element={element}
        zoom={zoom}
        selectedTablePart={selectedTablePart}
        selectedTableCell={selectedTableCell}
        editingTableCell={editingTableCell}
        onTableDividerPointerDown={onTableDividerPointerDown}
        onTableCellDoubleClick={onTableCellDoubleClick}
        onTableCellChange={onTableCellChange}
        onTableCellEditEnd={onTableCellEditEnd}
      />
      {selected && (
        <button
          className={element.locked ? "lock-handle locked" : "lock-handle"}
          aria-label={element.locked ? "Odblokuj element" : "Zablokuj element"}
          title={element.locked ? "Odblokuj element" : "Zablokuj element"}
          onPointerDown={(event) => event.stopPropagation()}
          onClick={(event) => {
            event.stopPropagation();
            onToggleLock();
          }}
        >
          {element.locked ? <Lock size={13} /> : <Unlock size={13} />}
        </button>
      )}
      {selected && !element.locked && (
        <button className="resize-handle" aria-label="Zmień rozmiar" onPointerDown={onResizeStart} />
      )}
    </div>
  );
}

function ElementContent({
  element,
  zoom,
  selectedTablePart,
  selectedTableCell,
  editingTableCell,
  onTableDividerPointerDown,
  onTableCellDoubleClick,
  onTableCellChange,
  onTableCellEditEnd,
}: {
  element: DocumentElement;
  zoom: number;
  selectedTablePart: TableSelection | null;
  selectedTableCell: TableCellSelection | null;
  editingTableCell: TableCellSelection | null;
  onTableDividerPointerDown: (
    event: PointerEvent<HTMLButtonElement>,
    table: TableElement,
    axis: "column" | "row",
    index: number,
  ) => void;
  onTableCellDoubleClick: (
    event: MouseEvent<HTMLSpanElement>,
    table: TableElement,
    row: number,
    column: number,
  ) => void;
  onTableCellChange: (table: TableElement, row: number, column: number, text: string) => void;
  onTableCellEditEnd: () => void;
}) {
  if (element.kind === "text") {
    return (
      <div
        className="text-element"
        style={{
          fontSize: element.fontSize * zoom,
          fontWeight: element.fontWeight,
          textAlign: element.align,
          color: element.color ?? "#172033",
          writingMode: element.orientation === "vertical" ? "vertical-rl" : "horizontal-tb",
          transform:
            element.orientation === "vertical" && element.verticalDirection === "counterclockwise"
              ? "rotate(180deg)"
              : "none",
        }}
      >
        {element.text}
      </div>
    );
  }

  if (element.kind === "marker") {
    return (
      <div
        className="marker-element"
        style={{
          fontSize: element.fontSize * zoom,
          fontWeight: element.fontWeight,
        }}
      >
        {element.code}
      </div>
    );
  }

  if (element.kind === "box") {
    return <div className="box-element" style={{ borderWidth: element.borderWidth * zoom, background: element.fill }} />;
  }

  if (element.kind === "line") {
    return (
      <div className={element.orientation === "vertical" ? "line-element vertical" : "line-element horizontal"}>
        <span
          style={
            element.orientation === "vertical"
              ? { width: element.borderWidth * zoom }
              : { height: element.borderWidth * zoom }
          }
        />
      </div>
    );
  }

  if (element.kind === "table") {
    const table = normalizeTable(element);
    const columnWidths = table.columnWidthsMm ?? [];
    const rowHeights = table.rowHeightsMm ?? [];

    return (
      <div className="table-shell">
        <div
          className="table-element"
          style={{
            gridTemplateColumns: columnWidths.map((width) => `${width * PX_PER_MM * zoom}px`).join(" "),
            gridTemplateRows: rowHeights.map((height) => `${height * PX_PER_MM * zoom}px`).join(" "),
            fontSize: table.fontSize * zoom,
            "--table-border-width": `${(table.borderWidth ?? 1) * zoom}px`,
            "--table-cell-padding-y": `${2 * zoom}px`,
            "--table-cell-padding-x": `${4 * zoom}px`,
          } as CSSProperties}
        >
          {Array.from({ length: table.rows * table.columns }).map((_, index) => {
            const row = Math.floor(index / table.columns);
            const column = index % table.columns;
            const cell = table.cells?.[index] ?? createEmptyCell();
            const active = selectedTableCell?.tableId === table.id && selectedTableCell.row === row && selectedTableCell.column === column;
            const editing =
              editingTableCell?.tableId === table.id && editingTableCell.row === row && editingTableCell.column === column;

            return (
              <span
                key={index}
                className={active ? "table-cell active" : "table-cell"}
                style={{
                  justifyContent: alignToFlex(cell.align),
                  alignItems: verticalAlignToFlex(cell.verticalAlign),
                  textAlign: cell.align,
                }}
                onDoubleClick={(event) => onTableCellDoubleClick(event, table, row, column)}
              >
                {editing ? (
                  <CellTextarea
                    value={cell.text}
                    align={cell.align}
                    onChange={(text) => onTableCellChange(table, row, column, text)}
                    onEditEnd={onTableCellEditEnd}
                  />
                ) : (
                  cell.text
                )}
              </span>
            );
          })}
        </div>

        {columnWidths.slice(0, -1).map((_, index) => (
          <button
            key={`column-${index}`}
            className={
              selectedTablePart?.axis === "column" && selectedTablePart.index === index
                ? "table-divider column active"
                : "table-divider column"
            }
            style={{ left: cumulativeSize(columnWidths, index + 1) * PX_PER_MM * zoom }}
            aria-label={`Zmień szerokość kolumny ${index + 1}`}
            onPointerDown={(event) => onTableDividerPointerDown(event, table, "column", index)}
          />
        ))}

        {rowHeights.slice(0, -1).map((_, index) => (
          <button
            key={`row-${index}`}
            className={
              selectedTablePart?.axis === "row" && selectedTablePart.index === index
                ? "table-divider row active"
                : "table-divider row"
            }
            style={{ top: cumulativeSize(rowHeights, index + 1) * PX_PER_MM * zoom }}
            aria-label={`Zmień wysokość wiersza ${index + 1}`}
            onPointerDown={(event) => onTableDividerPointerDown(event, table, "row", index)}
          />
        ))}
      </div>
    );
  }

  if (element.kind === "image") {
    return element.src ? <img src={element.src} alt="" draggable={false} /> : <div className="image-placeholder">Logo / obraz</div>;
  }

  if (element.kind === "arrow") {
    return <ArrowContent arrow={element} zoom={zoom} />;
  }

  return (
    <div
      className={isTwoDimensionalBarcode(element.symbology) ? "barcode-element two-dimensional" : "barcode-element"}
      style={
        {
          "--barcode-gap": `${(isTwoDimensionalBarcode(element.symbology) ? 4 : 3) * zoom}px`,
          "--barcode-stripe-gap": `${zoom}px`,
          "--barcode-stripe-min-width": `${zoom}px`,
          "--barcode-label-font-size": `${9 * zoom}px`,
          "--barcode-matrix-max-size": `${72 * zoom}px`,
        } as CSSProperties
      }
    >
      {isTwoDimensionalBarcode(element.symbology) ? (
        <div className={`barcode-matrix ${element.symbology}`} aria-hidden="true">
          {createMatrixPattern(element.value, element.symbology).map((cell, index) => (
            <span key={index} className={cell} />
          ))}
        </div>
      ) : (
        <div className="barcode-lines" aria-hidden="true">
          {createBarcodePattern(element.value).map((width, index) => (
            <span
              key={`${width}-${index}`}
              className={index % 2 === 0 ? "barcode-stripe dark" : "barcode-stripe light"}
              style={{ flexGrow: width }}
            />
          ))}
        </div>
      )}
      <span>
        {barcodeSymbologyLabel(element.symbology)} · {element.value}
      </span>
    </div>
  );
}

function PropertiesPanel({
  element,
  unit,
  selectedTablePart,
  selectedTableCell,
  onChange,
  onImageUpload,
}: {
  element: DocumentElement;
  unit: MeasurementUnit;
  selectedTablePart: TableSelection | null;
  selectedTableCell: TableCellSelection | null;
  onChange: (patch: Partial<DocumentElement>) => void;
  onImageUpload: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <div className="properties-grid">
      <label>
        Nazwa
        <input value={element.name} onChange={(event) => onChange({ name: event.target.value })} />
      </label>
      <NumberField label={`X ${units[unit].suffix}`} valueMm={element.xMm} unit={unit} onChange={(xMm) => onChange({ xMm })} />
      <NumberField label={`Y ${units[unit].suffix}`} valueMm={element.yMm} unit={unit} onChange={(yMm) => onChange({ yMm })} />
      <NumberField
        label={`Szer. ${units[unit].suffix}`}
        valueMm={element.widthMm}
        unit={unit}
        minMm={1}
        onChange={(widthMm) => onChange({ widthMm })}
      />
      {element.kind !== "line" && (
        <NumberField
          label={`Wys. ${units[unit].suffix}`}
          valueMm={element.heightMm}
          unit={unit}
          minMm={0}
          onChange={(heightMm) => onChange({ heightMm })}
        />
      )}

      {element.kind === "line" && (
        <LineProperties line={element} onChange={onChange} />
      )}

      {element.kind === "box" && (
        <NumberInput
          label="Grubość kreski"
          min={0}
          step={0.1}
          value={element.borderWidth}
          onChange={(borderWidth) => onChange({ borderWidth } as Partial<DocumentElement>)}
        />
      )}

      {element.kind === "arrow" && (
        <ArrowProperties arrow={element} unit={unit} onChange={onChange} />
      )}

      {element.kind === "marker" && (
        <>
          <label>
            Kod
            <input value={element.code} onChange={(event) => onChange({ code: event.target.value })} />
          </label>
          <label>
            Font
            <input
              type="number"
              min="6"
              value={element.fontSize}
              onChange={(event) => onChange({ fontSize: Number(event.target.value) })}
            />
          </label>
        </>
      )}

      {element.kind === "text" && (
        <>
          <label className="wide">
            Tekst
            <textarea value={element.text} onChange={(event) => onChange({ text: event.target.value })} />
          </label>
          <label>
            Font
            <input
              type="number"
              value={element.fontSize}
              onChange={(event) => onChange({ fontSize: Number(event.target.value) })}
            />
          </label>
          <label>
            Waga
            <select
              value={element.fontWeight}
              onChange={(event) => onChange({ fontWeight: event.target.value as "400" | "700" })}
            >
              <option value="400">Normal</option>
              <option value="700">Bold</option>
            </select>
          </label>
          <TextColorPicker color={element.color ?? "#172033"} onChange={(color) => onChange({ color })} />
          <label>
            Orientacja
            <select
              value={element.orientation ?? "horizontal"}
              onChange={(event) => onChange({ orientation: event.target.value as "horizontal" | "vertical" })}
            >
              <option value="horizontal">Pozioma</option>
              <option value="vertical">Pionowa</option>
            </select>
          </label>
          {(element.orientation ?? "horizontal") === "vertical" && (
            <label>
              Kierunek
              <select
                value={element.verticalDirection ?? "clockwise"}
                onChange={(event) =>
                  onChange({ verticalDirection: event.target.value as "clockwise" | "counterclockwise" })
                }
              >
                <option value="clockwise">W prawo</option>
                <option value="counterclockwise">W lewo</option>
              </select>
            </label>
          )}
        </>
      )}

      {element.kind === "table" && (
        <TableProperties
          table={element}
          unit={unit}
          selectedTablePart={selectedTablePart?.tableId === element.id ? selectedTablePart : null}
          selectedTableCell={selectedTableCell?.tableId === element.id ? selectedTableCell : null}
          onChange={onChange}
        />
      )}

      {element.kind === "barcode" && (
        <>
          <label className="wide">
            Typ kodu
            <select
              value={element.symbology}
              onChange={(event) =>
                onChange({
                  symbology: event.target.value as Extract<DocumentElement, { kind: "barcode" }>["symbology"],
                })
              }
            >
              <optgroup label="Jednowymiarowe">
                {barcodeSymbologies
                  .filter((symbology) => symbology.group === "1D")
                  .map((symbology) => (
                    <option key={symbology.value} value={symbology.value}>
                      {symbology.label}
                    </option>
                  ))}
              </optgroup>
              <optgroup label="Dwuwymiarowe">
                {barcodeSymbologies
                  .filter((symbology) => symbology.group === "2D")
                  .map((symbology) => (
                    <option key={symbology.value} value={symbology.value}>
                      {symbology.label}
                    </option>
                  ))}
              </optgroup>
            </select>
          </label>
          <label className="wide">
            Wartość kodu
            <input value={element.value} onChange={(event) => onChange({ value: event.target.value })} />
          </label>
        </>
      )}

      {element.kind === "image" && (
        <label className="wide">
          Plik obrazu
          <input type="file" accept="image/*" onChange={onImageUpload} />
        </label>
      )}
    </div>
  );
}

function TableProperties({
  table,
  unit,
  selectedTablePart,
  selectedTableCell,
  onChange,
}: {
  table: TableElement;
  unit: MeasurementUnit;
  selectedTablePart: TableSelection | null;
  selectedTableCell: TableCellSelection | null;
  onChange: (patch: Partial<DocumentElement>) => void;
}) {
  const normalizedTable = normalizeTable(table);
  const columnWidths = normalizedTable.columnWidthsMm ?? [];
  const rowHeights = normalizedTable.rowHeightsMm ?? [];
  const selectedCellIndex =
    selectedTableCell && selectedTableCell.row < table.rows && selectedTableCell.column < table.columns
      ? selectedTableCell.row * table.columns + selectedTableCell.column
      : null;
  const selectedCell = typeof selectedCellIndex === "number" ? normalizedTable.cells?.[selectedCellIndex] : null;

  function changeColumnWidth(index: number, widthMm: number) {
    const columnWidthsMm = [...columnWidths];
    columnWidthsMm[index] = widthMm;
    onChange({ columnWidthsMm, widthMm: roundMm(sum(columnWidthsMm)) } as Partial<DocumentElement>);
  }

  function changeRowHeight(index: number, heightMm: number) {
    const rowHeightsMm = [...rowHeights];
    rowHeightsMm[index] = heightMm;
    onChange({ rowHeightsMm, heightMm: roundMm(sum(rowHeightsMm)) } as Partial<DocumentElement>);
  }

  function changeCell(patch: Partial<NonNullable<TableElement["cells"]>[number]>) {
    if (typeof selectedCellIndex !== "number") {
      return;
    }

    const cells = [...(normalizedTable.cells ?? [])];
    cells[selectedCellIndex] = {
      ...createEmptyCell(),
      ...cells[selectedCellIndex],
      ...patch,
    };
    onChange({ cells } as Partial<DocumentElement>);
  }

  return (
    <>
      <label>
        Wiersze
        <input
          type="number"
          min="1"
          value={table.rows}
            onChange={(event) => {
              const rows = Math.max(1, Number(event.target.value));
              onChange({
                rows,
                rowHeightsMm: resizeDistribution(rowHeights, rows, table.heightMm),
                cells: resizeTableCells(normalizedTable.cells ?? [], table.rows, table.columns, rows, table.columns),
              } as Partial<DocumentElement>);
            }}
        />
      </label>
      <label>
        Kolumny
        <input
          type="number"
          min="1"
          value={table.columns}
            onChange={(event) => {
              const columns = Math.max(1, Number(event.target.value));
              onChange({
                columns,
                columnWidthsMm: resizeDistribution(columnWidths, columns, table.widthMm),
                cells: resizeTableCells(normalizedTable.cells ?? [], table.rows, table.columns, table.rows, columns),
              } as Partial<DocumentElement>);
          }}
        />
      </label>
      <label>
        Font
        <input
          type="number"
          min="6"
          value={table.fontSize}
          onChange={(event) => onChange({ fontSize: Number(event.target.value) } as Partial<DocumentElement>)}
        />
      </label>
      <NumberInput
        label="Grubość kreski"
        min={0}
        step={0.1}
        value={normalizedTable.borderWidth ?? 1}
        onChange={(borderWidth) => onChange({ borderWidth } as Partial<DocumentElement>)}
      />

      {selectedCell && selectedTableCell && (
        <div className="selected-table-part wide">
          <strong>
            Komórka {selectedTableCell.row + 1}:{selectedTableCell.column + 1}
          </strong>
          <label>
            Treść
            <textarea value={selectedCell.text} onChange={(event) => changeCell({ text: event.target.value })} />
          </label>
          <label>
            Poziomo
            <select
              value={selectedCell.align}
              onChange={(event) => changeCell({ align: event.target.value as "left" | "center" | "right" })}
            >
              <option value="left">Lewa</option>
              <option value="center">Środek</option>
              <option value="right">Prawa</option>
            </select>
          </label>
          <label>
            Pionowo
            <select
              value={selectedCell.verticalAlign}
              onChange={(event) => changeCell({ verticalAlign: event.target.value as "top" | "middle" | "bottom" })}
            >
              <option value="top">Góra</option>
              <option value="middle">Środek</option>
              <option value="bottom">Dół</option>
            </select>
          </label>
        </div>
      )}

      {selectedTablePart && (
        <div className="selected-table-part wide">
          <strong>
            {selectedTablePart.axis === "column"
              ? `Kolumna ${selectedTablePart.index + 1}`
              : `Wiersz ${selectedTablePart.index + 1}`}
          </strong>
          <NumberField
            label={
              selectedTablePart.axis === "column"
                ? `Szerokość ${units[unit].suffix}`
                : `Wysokość ${units[unit].suffix}`
            }
            valueMm={
              selectedTablePart.axis === "column"
                ? columnWidths[selectedTablePart.index]
                : rowHeights[selectedTablePart.index]
            }
            unit={unit}
            minMm={selectedTablePart.axis === "column" ? 5 : 4}
            onChange={(valueMm) => {
              if (selectedTablePart.axis === "column") {
                changeColumnWidth(selectedTablePart.index, valueMm);
              } else {
                changeRowHeight(selectedTablePart.index, valueMm);
              }
            }}
          />
        </div>
      )}

      <div className="table-size-list wide">
        <strong>Kolumny</strong>
        {columnWidths.map((widthMm, index) => (
          <NumberField
            key={`column-width-${index}`}
            label={`${index + 1}`}
            valueMm={widthMm}
            unit={unit}
            minMm={5}
            onChange={(valueMm) => changeColumnWidth(index, valueMm)}
          />
        ))}
      </div>

      <div className="table-size-list wide">
        <strong>Wiersze</strong>
        {rowHeights.map((heightMm, index) => (
          <NumberField
            key={`row-height-${index}`}
            label={`${index + 1}`}
            valueMm={heightMm}
            unit={unit}
            minMm={4}
            onChange={(valueMm) => changeRowHeight(index, valueMm)}
          />
        ))}
      </div>
    </>
  );
}

function TextColorPicker({ color, onChange }: { color: string; onChange: (color: string) => void }) {
  const presets = [
    "#172033",
    "#000000",
    "#475467",
    "#b42318",
    "#b54708",
    "#027a48",
    "#175cd3",
    "#6941c6",
  ];

  return (
    <div className="text-color-picker wide">
      <strong>Kolor tekstu</strong>
      <div className="color-preset-grid">
        {presets.map((preset) => (
          <button
            key={preset}
            className={color.toLowerCase() === preset.toLowerCase() ? "color-preset active" : "color-preset"}
            style={{ background: preset }}
            type="button"
            aria-label={`Ustaw kolor ${preset}`}
            onClick={() => onChange(preset)}
          />
        ))}
      </div>
      <label>
        Własny kolor
        <input type="color" value={color} onChange={(event) => onChange(event.target.value)} />
      </label>
    </div>
  );
}

function MarkerPanel({
  panelRef,
  markers,
  columns,
  columnWidths,
  selectedMarkerId,
  open,
  onOpenChange,
  onSelectMarker,
  onChangeMarker,
  onCommitMarkerEdit,
  onRenameColumn,
  onAddColumn,
  onDeleteColumn,
  onCopyMarkers,
  onColumnWidthChange,
}: {
  panelRef: RefObject<HTMLElement | null>;
  markers: Array<Extract<DocumentElement, { kind: "marker" }>>;
  columns: MarkerColumn[];
  columnWidths: MarkerColumnWidths;
  selectedMarkerId: string | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSelectMarker: (id: string) => void;
  onChangeMarker: (id: string, patch: Partial<DocumentElement>) => void;
  onCommitMarkerEdit: () => void;
  onRenameColumn: (columnId: string, label: string) => void;
  onAddColumn: () => void;
  onDeleteColumn: (columnId: string) => void;
  onCopyMarkers: () => void;
  onColumnWidthChange: (column: MarkerColumnKey, width: number) => void;
}) {
  const [editingColumnId, setEditingColumnId] = useState<string | null>(null);
  const totalColumnWidth = columns.reduce((total, column) => total + markerColumnWidth(columnWidths, column.id), 0);
  const selectedRowRef = useRef<HTMLTableRowElement | null>(null);
  const resizingColumnRef = useRef<{ column: MarkerColumnKey; startX: number; startWidth: number } | null>(null);

  useEffect(() => {
    if (!open || !selectedMarkerId) {
      return;
    }

    window.setTimeout(() => {
      selectedRowRef.current?.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 80);
  }, [open, selectedMarkerId]);

  function onColumnResizeMove(event: PointerEvent<HTMLDivElement>) {
    const resize = resizingColumnRef.current;
    if (!resize) {
      return;
    }

    onColumnWidthChange(resize.column, resize.startWidth + event.clientX - resize.startX);
  }

  function stopColumnResize() {
    resizingColumnRef.current = null;
  }

  return (
    <>
      <button className={`marker-panel-toggle ${open ? "open" : ""}`} type="button" onClick={() => onOpenChange(!open)}>
        {open ? <ChevronDown size={16} /> : <ChevronUp size={16} />}
        {open ? "Zwiń panel" : "Pokaż znaczniki"}
      </button>
      <section
      ref={panelRef}
      className={`marker-panel ${open ? "open" : ""}`}
      onPointerMove={onColumnResizeMove}
      onPointerUp={stopColumnResize}
      onPointerLeave={stopColumnResize}
    >
      <header className="marker-panel-header">
        <div>
          <strong>Znaczniki dokumentu</strong>
          <span>{markers.length} pozycji</span>
        </div>
        <div className="marker-panel-actions">
          <button type="button" onClick={onAddColumn}>
            + Kolumna
          </button>
          <button type="button" onClick={onCopyMarkers}>
            Kopiuj tabelę
          </button>
        </div>
      </header>
      <div className="marker-table-wrap">
        <table className="marker-table" style={{ width: totalColumnWidth }}>
          <colgroup>
            {columns.map((column) => (
              <col key={column.id} style={{ width: markerColumnWidth(columnWidths, column.id) }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              {columns.map((column) => (
                <th key={column.id}>
                  <div className="marker-column-heading">
                    {editingColumnId === column.id ? (
                      <input
                        value={column.label}
                        autoFocus
                        onChange={(event) => onRenameColumn(column.id, event.target.value)}
                        onBlur={() => setEditingColumnId(null)}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === "Escape") {
                            setEditingColumnId(null);
                          }
                        }}
                      />
                    ) : (
                      <span>{column.label}</span>
                    )}
                    <button
                      className="marker-heading-button"
                      type="button"
                      aria-label={`Edytuj nazwę kolumny ${column.label}`}
                      onClick={() => setEditingColumnId(column.id)}
                    >
                      <Pencil size={12} />
                    </button>
                    {columns.length > 1 && (
                      <button
                        className="marker-heading-button danger"
                        type="button"
                        aria-label={`Usuń kolumnę ${column.label}`}
                        onClick={() => onDeleteColumn(column.id)}
                      >
                        <Trash2 size={12} />
                      </button>
                    )}
                  </div>
                  <button
                    className="marker-column-resizer"
                    type="button"
                    aria-label={`Zmień szerokość kolumny ${column.label}`}
                    onPointerDown={(event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      resizingColumnRef.current = {
                        column: column.id,
                        startX: event.clientX,
                        startWidth: markerColumnWidth(columnWidths, column.id),
                      };
                    }}
                  />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {markers.map((marker) => (
              <tr
                key={marker.id}
                ref={marker.id === selectedMarkerId ? selectedRowRef : null}
                className={marker.id === selectedMarkerId ? "active" : ""}
                onClick={() => onSelectMarker(marker.id)}
              >
                {columns.map((column) => (
                  <td key={column.id}>
                    {column.field === "managementRules" ? (
                      <AutosizeTextarea
                        value={markerColumnValue(marker, column)}
                        onChange={(event) => onChangeMarker(marker.id, markerColumnPatch(marker, column, event.target.value))}
                        onBlur={onCommitMarkerEdit}
                      />
                    ) : (
                      <input
                        value={markerColumnValue(marker, column)}
                        onChange={(event) => onChangeMarker(marker.id, markerColumnPatch(marker, column, event.target.value))}
                        onBlur={onCommitMarkerEdit}
                      />
                    )}
                  </td>
                ))}
              </tr>
            ))}
            {markers.length === 0 && (
              <tr>
                <td colSpan={columns.length}>Dodaj znacznik z menu kontekstowego na kanwie.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      </section>
    </>
  );
}

function ArrowContent({ arrow, zoom }: { arrow: Extract<DocumentElement, { kind: "arrow" }>; zoom: number }) {
  const isVertical = arrow.orientation === "vertical";
  const lengthMm = isVertical ? arrow.heightMm : arrow.widthMm;
  const label = arrow.labelText?.trim() ? arrow.labelText : formatMeasurement(lengthMm, arrow.labelUnit);
  const headSize = arrow.headSize ?? Math.max(5, arrow.strokeWidth * 3);
  const labelPosition = Math.min(95, Math.max(5, arrow.labelPosition));
  const viewBox = isVertical ? `0 0 ${Math.max(24, arrow.widthMm)} ${arrow.heightMm}` : `0 0 ${arrow.widthMm} ${Math.max(24, arrow.heightMm)}`;
  const centerX = isVertical ? Math.max(24, arrow.widthMm) / 2 : arrow.widthMm / 2;
  const centerY = isVertical ? arrow.heightMm / 2 : Math.max(24, arrow.heightMm) / 2;
  const start = headSize / 2;
  const end = lengthMm - headSize / 2;
  const labelX = isVertical ? centerX : (arrow.widthMm * labelPosition) / 100;
  const labelY = isVertical ? (arrow.heightMm * labelPosition) / 100 : centerY;
  const labelOffset = 11 + arrow.strokeWidth;
  const adjustedLabelX = isVertical
    ? labelX + (arrow.labelSide === "right" ? labelOffset : arrow.labelSide === "left" ? -labelOffset : 0)
    : labelX;
  const adjustedLabelY = isVertical
    ? labelY
    : labelY + (arrow.labelSide === "below" ? labelOffset : arrow.labelSide === "above" ? -labelOffset : 0);
  const labelWidth = isVertical ? Math.max(24, arrow.widthMm) : arrow.widthMm;
  const labelHeight = isVertical ? arrow.heightMm : Math.max(24, arrow.heightMm);

  return (
    <div className="arrow-element">
      <svg viewBox={viewBox} preserveAspectRatio="none">
        {isVertical ? (
          <line x1={centerX} y1={start} x2={centerX} y2={end} strokeWidth={arrow.strokeWidth} />
        ) : (
          <line x1={start} y1={centerY} x2={end} y2={centerY} strokeWidth={arrow.strokeWidth} />
        )}
        {(arrow.head === "start" || arrow.head === "both") && (
          <ArrowHead
            x={isVertical ? centerX : start}
            y={isVertical ? start : centerY}
            direction={isVertical ? "up" : "left"}
            size={headSize}
            style={arrow.headStyle}
            strokeWidth={arrow.strokeWidth}
          />
        )}
        {(arrow.head === "end" || arrow.head === "both") && (
          <ArrowHead
            x={isVertical ? centerX : end}
            y={isVertical ? end : centerY}
            direction={isVertical ? "down" : "right"}
            size={headSize}
            style={arrow.headStyle}
            strokeWidth={arrow.strokeWidth}
          />
        )}
      </svg>
      {arrow.showLabel && (
        <span
          className={isVertical ? "arrow-label vertical" : "arrow-label"}
          style={{
            left: `${(adjustedLabelX / labelWidth) * 100}%`,
            top: `${(adjustedLabelY / labelHeight) * 100}%`,
            fontSize: (arrow.labelFontSize ?? 7) * zoom,
          }}
        >
          {label}
        </span>
      )}
    </div>
  );
}

function ArrowProperties({
  arrow,
  unit,
  onChange,
}: {
  arrow: Extract<DocumentElement, { kind: "arrow" }>;
  unit: MeasurementUnit;
  onChange: (patch: Partial<DocumentElement>) => void;
}) {
  const lengthMm = arrow.orientation === "vertical" ? arrow.heightMm : arrow.widthMm;

  function changeOrientation(orientation: "horizontal" | "vertical") {
    if (orientation === arrow.orientation) {
      return;
    }

    onChange({
      orientation,
      widthMm: arrow.heightMm,
      heightMm: arrow.widthMm,
      labelSide: orientation === "vertical" ? "right" : "above",
    } as Partial<DocumentElement>);
  }

  return (
    <div className="arrow-properties wide">
      <label>
        Orientacja
        <select value={arrow.orientation} onChange={(event) => changeOrientation(event.target.value as "horizontal" | "vertical")}>
          <option value="horizontal">Pozioma</option>
          <option value="vertical">Pionowa</option>
        </select>
      </label>
      <label>
        Długość {units[unit].suffix}
        <input
          type="number"
          step={units[unit].step}
          value={formatNumber(fromMm(lengthMm, unit))}
          onChange={(event) => {
            const length = Math.max(5, toMm(Number(event.target.value), unit));
            onChange(arrow.orientation === "vertical" ? { heightMm: length } : { widthMm: length });
          }}
        />
      </label>
      <label>
        Groty
        <select value={arrow.head} onChange={(event) => onChange({ head: event.target.value as "end" | "start" | "both" })}>
          <option value="end">Na końcu</option>
          <option value="start">Na początku</option>
          <option value="both">Z obu stron</option>
        </select>
      </label>
      <label>
        Styl grotu
        <select value={arrow.headStyle} onChange={(event) => onChange({ headStyle: event.target.value as "triangle" | "bar" | "open" })}>
          <option value="triangle">Pełny</option>
          <option value="open">Otwarty</option>
          <option value="bar">Kreska</option>
        </select>
      </label>
      <label>
        Grubość px
        <input
          type="number"
          min="0.5"
          max="20"
          step="0.5"
          value={arrow.strokeWidth}
          onChange={(event) => onChange({ strokeWidth: Math.max(0.5, Number(event.target.value)) })}
        />
      </label>
      <label>
        Rozmiar grotu
        <input
          type="number"
          min="3"
          max="20"
          value={arrow.headSize ?? 5}
          onChange={(event) => onChange({ headSize: Math.max(3, Number(event.target.value)) })}
        />
      </label>
      <label className="checkbox-row wide">
        <input type="checkbox" checked={arrow.showLabel} onChange={(event) => onChange({ showLabel: event.target.checked })} />
        <span>Pokaż label długości</span>
      </label>
      {arrow.showLabel && (
        <>
          <label>
            Jednostka labelki
            <select value={arrow.labelUnit} onChange={(event) => onChange({ labelUnit: event.target.value as MeasurementUnit })}>
              <option value="mm">mm</option>
              <option value="cm">cm</option>
              <option value="in">cale</option>
            </select>
          </label>
          <label>
            Font labelki
            <input
              type="number"
              min="4"
              max="24"
              value={arrow.labelFontSize ?? 7}
              onChange={(event) => onChange({ labelFontSize: Math.max(4, Number(event.target.value)) })}
            />
          </label>
          <label className="wide">
            Własny tekst labelki
            <input
              value={arrow.labelText ?? ""}
              placeholder={formatMeasurement(lengthMm, arrow.labelUnit)}
              onChange={(event) => onChange({ labelText: event.target.value })}
            />
          </label>
          <label>
            Strona labelki
            <select value={arrow.labelSide} onChange={(event) => onChange({ labelSide: event.target.value as "above" | "below" | "left" | "right" })}>
              {arrow.orientation === "vertical" ? (
                <>
                  <option value="left">Lewa</option>
                  <option value="right">Prawa</option>
                </>
              ) : (
                <>
                  <option value="above">Nad</option>
                  <option value="below">Pod</option>
                </>
              )}
            </select>
          </label>
          <label className="wide">
            Pozycja labelki
            <input type="range" min="5" max="95" value={arrow.labelPosition} onChange={(event) => onChange({ labelPosition: Number(event.target.value) })} />
          </label>
        </>
      )}
    </div>
  );
}

function ArrowHead({
  x,
  y,
  direction,
  size,
  style,
  strokeWidth,
}: {
  x: number;
  y: number;
  direction: "left" | "right" | "up" | "down";
  size: number;
  style: "triangle" | "bar" | "open";
  strokeWidth: number;
}) {
  const angle = direction === "right" ? 0 : direction === "down" ? 90 : direction === "left" ? 180 : 270;

  if (style === "bar") {
    return <line className="arrow-head" x1={x} y1={y - size / 2} x2={x} y2={y + size / 2} strokeWidth={strokeWidth} transform={`rotate(${angle} ${x} ${y})`} />;
  }

  if (style === "open") {
    return (
      <path
        className="arrow-head"
        d={`M ${x - size} ${y - size / 3} L ${x} ${y} L ${x - size} ${y + size / 3}`}
        strokeWidth={strokeWidth}
        transform={`rotate(${angle} ${x} ${y})`}
      />
    );
  }

  return (
    <path
      className="arrow-head filled"
      d={`M ${x} ${y} L ${x - size} ${y - size / 3} L ${x - size} ${y + size / 3} Z`}
      transform={`rotate(${angle} ${x} ${y})`}
    />
  );
}

function LineProperties({
  line,
  onChange,
}: {
  line: Extract<DocumentElement, { kind: "line" }>;
  onChange: (patch: Partial<DocumentElement>) => void;
}) {
  const presets = [1, 2, 3, 4, 6];

  function changeOrientation(orientation: "horizontal" | "vertical") {
    const currentOrientation = line.orientation ?? "horizontal";
    if (orientation === currentOrientation) {
      return;
    }

    onChange({
      orientation,
      widthMm: line.heightMm,
      heightMm: line.widthMm,
    } as Partial<DocumentElement>);
  }

  return (
    <div className="line-presets wide">
      <label>
        Orientacja
        <select
          value={line.orientation ?? "horizontal"}
          onChange={(event) => changeOrientation(event.target.value as "horizontal" | "vertical")}
        >
          <option value="horizontal">Pozioma</option>
          <option value="vertical">Pionowa</option>
        </select>
      </label>
      <strong>Grubość linii</strong>
      <div className="line-preset-grid">
        {presets.map((preset) => (
          <button
            key={preset}
            className={line.borderWidth === preset ? "line-preset active" : "line-preset"}
            onClick={() => onChange({ borderWidth: preset })}
            type="button"
          >
            <span style={{ height: preset }} />
            {preset}px
          </button>
        ))}
      </div>
      <label>
        Własna grubość px
        <input
          type="number"
          min="1"
          max="20"
          value={line.borderWidth}
          onChange={(event) => onChange({ borderWidth: Math.max(1, Number(event.target.value)) })}
        />
      </label>
    </div>
  );
}

function CellTextarea({
  value,
  align,
  onChange,
  onEditEnd,
}: {
  value: string;
  align: "left" | "center" | "right";
  onChange: (text: string) => void;
  onEditEnd: () => void;
}) {
  const inputRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, []);

  function onKeyDown(event: ReactKeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === "Escape") {
      event.preventDefault();
      onEditEnd();
    }
  }

  function onBlur(_: FocusEvent<HTMLTextAreaElement>) {
    onEditEnd();
  }

  return (
    <textarea
      ref={inputRef}
      className="cell-editor"
      value={value}
      style={{ textAlign: align }}
      onChange={(event) => onChange(event.target.value)}
      onKeyDown={onKeyDown}
      onBlur={onBlur}
      onPointerDown={(event) => event.stopPropagation()}
      onDoubleClick={(event) => event.stopPropagation()}
    />
  );
}

function AutosizeTextarea({
  value,
  onChange,
  onBlur,
}: {
  value: string;
  onChange: (event: ChangeEvent<HTMLTextAreaElement>) => void;
  onBlur?: () => void;
}) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    const textarea = textareaRef.current;
    if (!textarea) {
      return;
    }

    textarea.style.height = "auto";
    textarea.style.height = `${textarea.scrollHeight}px`;
  }, [value]);

  return <textarea ref={textareaRef} value={value} onChange={onChange} onBlur={onBlur} />;
}

function NumberField({
  label,
  valueMm,
  unit,
  minMm,
  onChange,
}: {
  label: string;
  valueMm: number;
  unit: MeasurementUnit;
  minMm?: number;
  onChange: (valueMm: number) => void;
}) {
  return (
    <label>
      {label}
      <input
        type="number"
        step={units[unit].step}
        value={formatNumber(fromMm(valueMm, unit))}
        onChange={(event) => {
          const nextValue = toMm(Number(event.target.value), unit);
          onChange(typeof minMm === "number" ? Math.max(minMm, nextValue) : nextValue);
        }}
      />
    </label>
  );
}

function NumberInput({
  label,
  value,
  min,
  step = 1,
  onChange,
}: {
  label: string;
  value: number;
  min?: number;
  step?: number;
  onChange: (value: number) => void;
}) {
  return (
    <label>
      {label}
      <input
        type="number"
        min={min}
        step={step}
        value={formatNumber(value)}
        onChange={(event) => {
          const nextValue = Number(event.target.value);
          onChange(typeof min === "number" ? Math.max(min, nextValue) : nextValue);
        }}
      />
    </label>
  );
}

function roundMm(value: number) {
  return Math.round(value * 2) / 2;
}

function clampZoom(value: number) {
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Number(value.toFixed(2))));
}

function getPageOrientation(page: Pick<DocumentPage, "widthMm" | "heightMm">): PageOrientation {
  return page.widthMm > page.heightMm ? "landscape" : "portrait";
}

function orientPageSize(
  size: Pick<DocumentPage, "widthMm" | "heightMm">,
  orientation: PageOrientation,
): Pick<DocumentPage, "widthMm" | "heightMm"> {
  const shortSide = Math.min(size.widthMm, size.heightMm);
  const longSide = Math.max(size.widthMm, size.heightMm);

  return orientation === "landscape"
    ? { widthMm: longSide, heightMm: shortSide }
    : { widthMm: shortSide, heightMm: longSide };
}

function fromMm(valueMm: number, unit: MeasurementUnit) {
  return valueMm / units[unit].factorMm;
}

function toMm(value: number, unit: MeasurementUnit) {
  return roundMm(value * units[unit].factorMm);
}

function formatMeasurement(valueMm: number, unit: MeasurementUnit) {
  return `${formatNumber(fromMm(valueMm, unit))} ${units[unit].suffix}`;
}

function formatNumber(value: number) {
  return Number(value.toFixed(2));
}

function resolveAssistedPosition({
  element,
  elements,
  xMm,
  yMm,
  settings,
}: {
  element: DocumentElement;
  elements: DocumentElement[];
  xMm: number;
  yMm: number;
  settings: AssistSettings;
}) {
  let nextX = settings.stepMove ? snapToStep(xMm, settings.stepMm) : roundMm(xMm);
  let nextY = settings.stepMove ? snapToStep(yMm, settings.stepMm) : roundMm(yMm);
  let guides: ActiveGuides = { vertical: [], horizontal: [] };

  if (settings.snapToElements) {
    const snapped = snapToNeighborElements(element, elements, nextX, nextY);
    nextX = snapped.xMm;
    nextY = snapped.yMm;
    guides = snapped.guides;
  }

  return {
    xMm: nextX,
    yMm: nextY,
    guides,
  };
}

function resolveAssistedSize({
  element,
  elements,
  widthMm,
  heightMm,
  settings,
}: {
  element: DocumentElement;
  elements: DocumentElement[];
  widthMm: number;
  heightMm: number;
  settings: AssistSettings;
}) {
  const thresholdMm = 2;
  const steppedWidth = settings.stepMove ? snapToStep(widthMm, settings.stepMm) : widthMm;
  const steppedHeight = settings.stepMove ? snapToStep(heightMm, settings.stepMm) : heightMm;

  if (!settings.snapToElements) {
    return {
      widthMm: steppedWidth,
      heightMm: steppedHeight,
      guides: { vertical: [], horizontal: [] },
    };
  }

  const otherElements = elements.filter((item) => item.id !== element.id);
  const verticalTargets = otherElements.flatMap((item) => [
    item.xMm,
    item.xMm + item.widthMm / 2,
    item.xMm + item.widthMm,
  ]);
  const horizontalTargets = otherElements.flatMap((item) => [
    item.yMm,
    item.yMm + item.heightMm / 2,
    item.yMm + item.heightMm,
  ]);
  const rightEdge = element.xMm + steppedWidth;
  const bottomEdge = element.yMm + steppedHeight;
  const verticalTarget = findClosestValue(rightEdge, verticalTargets, thresholdMm);
  const horizontalTarget = findClosestValue(bottomEdge, horizontalTargets, thresholdMm);
  const snappedWidth = verticalTarget === null ? steppedWidth : Math.max(5, roundMm(verticalTarget - element.xMm));
  const snappedHeight = horizontalTarget === null ? steppedHeight : Math.max(1, roundMm(horizontalTarget - element.yMm));

  return {
    widthMm: snappedWidth,
    heightMm: snappedHeight,
    guides: {
      vertical: verticalTarget === null ? [] : [verticalTarget],
      horizontal: horizontalTarget === null ? [] : [horizontalTarget],
    },
  };
}

function snapToNeighborElements(element: DocumentElement, elements: DocumentElement[], xMm: number, yMm: number) {
  const thresholdMm = 2;
  const movingVerticalAnchors = [
    { key: "left", value: xMm, offset: 0 },
    { key: "center", value: xMm + element.widthMm / 2, offset: element.widthMm / 2 },
    { key: "right", value: xMm + element.widthMm, offset: element.widthMm },
  ];
  const movingHorizontalAnchors = [
    { key: "top", value: yMm, offset: 0 },
    { key: "middle", value: yMm + element.heightMm / 2, offset: element.heightMm / 2 },
    { key: "bottom", value: yMm + element.heightMm, offset: element.heightMm },
  ];
  const otherElements = elements.filter((item) => item.id !== element.id);
  const verticalTargets = otherElements.flatMap((item) => [
    item.xMm,
    item.xMm + item.widthMm / 2,
    item.xMm + item.widthMm,
  ]);
  const horizontalTargets = otherElements.flatMap((item) => [
    item.yMm,
    item.yMm + item.heightMm / 2,
    item.yMm + item.heightMm,
  ]);
  const verticalMatch = findClosestAnchor(movingVerticalAnchors, verticalTargets, thresholdMm);
  const horizontalMatch = findClosestAnchor(movingHorizontalAnchors, horizontalTargets, thresholdMm);

  return {
    xMm: verticalMatch ? roundMm(verticalMatch.target - verticalMatch.anchor.offset) : xMm,
    yMm: horizontalMatch ? roundMm(horizontalMatch.target - horizontalMatch.anchor.offset) : yMm,
    guides: {
      vertical: verticalMatch ? [verticalMatch.target] : [],
      horizontal: horizontalMatch ? [horizontalMatch.target] : [],
    },
  };
}

function findClosestAnchor(
  anchors: Array<{ key: string; value: number; offset: number }>,
  targets: number[],
  thresholdMm: number,
) {
  let bestMatch: { anchor: { key: string; value: number; offset: number }; target: number; distance: number } | null = null;

  for (const anchor of anchors) {
    for (const target of targets) {
      const distance = Math.abs(anchor.value - target);
      if (distance <= thresholdMm && (!bestMatch || distance < bestMatch.distance)) {
        bestMatch = { anchor, target, distance };
      }
    }
  }

  return bestMatch;
}

function findClosestValue(value: number, targets: number[], thresholdMm: number) {
  let bestMatch: { target: number; distance: number } | null = null;

  for (const target of targets) {
    const distance = Math.abs(value - target);
    if (distance <= thresholdMm && (!bestMatch || distance < bestMatch.distance)) {
      bestMatch = { target, distance };
    }
  }

  return bestMatch?.target ?? null;
}

function snapToStep(valueMm: number, stepMm: number) {
  return roundMm(Math.round(valueMm / stepMm) * stepMm);
}

function loadMarkerColumnWidths(): MarkerColumnWidths {
  try {
    const storedValue = localStorage.getItem("ldd-marker-column-widths");
    if (!storedValue) {
      return defaultMarkerColumnWidths;
    }

    return {
      ...defaultMarkerColumnWidths,
      ...(JSON.parse(storedValue) as Partial<MarkerColumnWidths>),
    };
  } catch {
    return defaultMarkerColumnWidths;
  }
}

function markerColumnWidth(widths: MarkerColumnWidths, columnId: string) {
  return widths[columnId] ?? 160;
}

function markerColumnValue(marker: Extract<DocumentElement, { kind: "marker" }>, column: MarkerColumn) {
  if (column.field === "custom") {
    return marker.customFields?.[column.customFieldId ?? column.id] ?? "";
  }

  return String(marker[column.field] ?? "");
}

function markerColumnPatch(
  marker: Extract<DocumentElement, { kind: "marker" }>,
  column: MarkerColumn,
  value: string,
): Partial<DocumentElement> {
  if (column.field === "custom") {
    const customFieldId = column.customFieldId ?? column.id;
    return {
      customFields: {
        ...marker.customFields,
        [customFieldId]: value,
      },
    } as Partial<DocumentElement>;
  }

  return { [column.field]: value } as Partial<DocumentElement>;
}

function addMarkerPagesToPdf(
  pdf: jsPDF,
  markers: Array<Extract<DocumentElement, { kind: "marker" }>>,
  markerColumns: MarkerColumn[],
  pageWidth: number,
  pageHeight: number,
) {
  if (markers.length === 0) {
    return;
  }

  const margin = 10;
  const headerHeight = 8;
  const minRowHeight = 9;
  const usableWidth = pageWidth - margin * 2;
  const columnWidth = usableWidth / markerColumns.length;
  const columnWidths = markerColumns.map(() => columnWidth);
  let y = margin;

  function addPageWithHeader() {
    pdf.addPage([pageWidth, pageHeight], pageWidth > pageHeight ? "landscape" : "portrait");
    y = margin;
    pdf.setFillColor(255, 255, 255);
    pdf.rect(0, 0, pageWidth, pageHeight, "F");
    pdf.setFont("helvetica", "bold");
    pdf.setTextColor(23, 32, 51);
    pdf.setFontSize(12);
    pdf.text("Tabela znaczników", margin, y);
    y += 8;
    drawTableHeader();
  }

  function drawTableHeader() {
    let x = margin;
    pdf.setDrawColor(210, 216, 226);
    pdf.setTextColor(23, 32, 51);
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(8);

    markerColumns.forEach((column, index) => {
      pdf.rect(x, y, columnWidths[index], headerHeight, "S");
      pdf.text(column.label, x + 1.5, y + 5.2, { maxWidth: columnWidths[index] - 3 });
      x += columnWidths[index];
    });
    y += headerHeight;
  }

  addPageWithHeader();
  pdf.setFont("helvetica", "normal");

  markers.forEach((marker) => {
    const values = markerColumns.map((column) => markerColumnValue(marker, column));
    const wrapped = values.map((value, index) => pdf.splitTextToSize(value, columnWidths[index] - 3) as string[]);
    const rowHeight = Math.max(minRowHeight, ...wrapped.map((lines) => lines.length * 4 + 4));

    if (y + rowHeight > pageHeight - margin) {
      addPageWithHeader();
      pdf.setFont("helvetica", "normal");
    }

    let x = margin;
    pdf.setDrawColor(210, 216, 226);
    pdf.setTextColor(23, 32, 51);
    pdf.setFontSize(8);
    wrapped.forEach((lines, index) => {
      pdf.rect(x, y, columnWidths[index], rowHeight);
      pdf.text(lines, x + 1.5, y + 5, { maxWidth: columnWidths[index] - 3 });
      x += columnWidths[index];
    });
    y += rowHeight;
  });
}

function sanitizeFileName(value: string) {
  return value.replace(/[<>:"/\\|?*]+/g, "-").trim() || "document";
}

function createDocxExport(
  documentImage: ArrayBuffer,
  documentName: string,
  markers: Array<Extract<DocumentElement, { kind: "marker" }>>,
  columns: MarkerColumn[],
) {
  const markerRows = markers.length > 0 ? markers : [];

  return new DocxDocument({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: documentName,
            heading: HeadingLevel.HEADING_1,
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                type: "png",
                data: documentImage,
                transformation: {
                  width: 560,
                  height: 792,
                },
              }),
            ],
          }),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({
            text: "Tabela znaczników",
            heading: HeadingLevel.HEADING_1,
          }),
          createDocxMarkerTable(markerRows, columns),
        ],
      },
    ],
  });
}

function createDocxMarkerTable(markers: Array<Extract<DocumentElement, { kind: "marker" }>>, columns: MarkerColumn[]) {
  const rows = [
    new TableRow({
      tableHeader: true,
      children: columns.map((column) =>
        new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: column.label, bold: true })] })],
        }),
      ),
    }),
    ...markers.map(
      (marker) =>
        new TableRow({
          children: columns.map(
            (column) =>
              new TableCell({
                children: [new Paragraph(markerColumnValue(marker, column))],
              }),
          ),
        }),
    ),
  ];

  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows,
  });
}

function markerTableToHtml(markers: Array<Extract<DocumentElement, { kind: "marker" }>>, columns: MarkerColumn[]) {
  const head = columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("");
  const body = markers
    .map(
      (marker) =>
        `<tr>${columns
          .map((column) => `<td>${escapeHtml(markerColumnValue(marker, column)).replace(/\n/g, "<br>")}</td>`)
          .join("")}</tr>`,
    )
    .join("");

  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function markerTableToPlainText(markers: Array<Extract<DocumentElement, { kind: "marker" }>>, columns: MarkerColumn[]) {
  return [
    columns.map((column) => column.label).join("\t"),
    ...markers.map((marker) => columns.map((column) => markerColumnValue(marker, column).replace(/\n/g, " ")).join("\t")),
  ].join("\n");
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function downloadBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function nextMarkerCode(code: string, existingCodes: string[]) {
  const match = code.match(/^([A-Za-z]*)(\d+)$/);

  if (!match) {
    return `${code} copy`;
  }

  const [, prefix, numericPart] = match;
  const padding = numericPart.length;
  let nextNumber = Number(numericPart) + 1;
  let nextCode = `${prefix}${String(nextNumber).padStart(padding, "0")}`;

  while (existingCodes.includes(nextCode)) {
    nextNumber += 1;
    nextCode = `${prefix}${String(nextNumber).padStart(padding, "0")}`;
  }

  return nextCode;
}

function createBarcodePattern(value: string) {
  const seed = Array.from(value).reduce((total, char) => total + char.charCodeAt(0), 0);
  return Array.from({ length: 72 }, (_, index) => {
    const mixed = (seed + index * 17 + (index % 7) * 11) % 5;
    return mixed + 1;
  });
}

function createMatrixPattern(
  value: string,
  symbology: Extract<DocumentElement, { kind: "barcode" }>["symbology"],
) {
  const seed = Array.from(value).reduce((total, char) => total + char.charCodeAt(0), 0);
  const size = 13;

  return Array.from({ length: size * size }, (_, index) => {
    const row = Math.floor(index / size);
    const column = index % size;
    const inTopLeft = row < 5 && column < 5;
    const inTopRight = row < 5 && column >= size - 5;
    const inBottomLeft = row >= size - 5 && column < 5;

    if (symbology === "qr" && (inTopLeft || inTopRight || inBottomLeft)) {
      const localRow = row < 5 ? row : row - (size - 5);
      const localColumn = column < 5 ? column : column - (size - 5);
      const border = localRow === 0 || localRow === 4 || localColumn === 0 || localColumn === 4;
      const center = localRow >= 2 && localRow <= 2 && localColumn >= 2 && localColumn <= 2;
      return border || center ? "filled finder" : "empty finder";
    }

    if (symbology === "datamatrix" && (row === 0 || column === 0 || row === size - 1 || column === size - 1)) {
      return "filled";
    }

    return (seed + row * 7 + column * 13) % 3 === 0 ? "filled" : "empty";
  });
}

function isTwoDimensionalBarcode(symbology: Extract<DocumentElement, { kind: "barcode" }>["symbology"]) {
  return symbology === "qr" || symbology === "datamatrix" || symbology === "pdf417";
}

function barcodeSymbologyLabel(symbology: Extract<DocumentElement, { kind: "barcode" }>["symbology"]) {
  return barcodeSymbologies.find((item) => item.value === symbology)?.label ?? symbology;
}

function normalizeTable(table: TableElement): TableElement {
  return {
    ...table,
    borderWidth: table.borderWidth ?? 1,
    columnWidthsMm: normalizeDistribution(table.columnWidthsMm, table.columns, table.widthMm),
    rowHeightsMm: normalizeDistribution(table.rowHeightsMm, table.rows, table.heightMm),
    cells: normalizeCells(table.cells, table.rows, table.columns),
  };
}

function createEmptyCell() {
  return {
    text: "",
    align: "left" as const,
    verticalAlign: "top" as const,
  };
}

function normalizeCells(cells: TableElement["cells"] | undefined, rows: number, columns: number) {
  return Array.from({ length: rows * columns }, (_, index) => ({
    ...createEmptyCell(),
    ...cells?.[index],
  }));
}

function resizeTableCells(
  cells: NonNullable<TableElement["cells"]>,
  oldRows: number,
  oldColumns: number,
  newRows: number,
  newColumns: number,
) {
  return Array.from({ length: newRows * newColumns }, (_, index) => {
    const row = Math.floor(index / newColumns);
    const column = index % newColumns;
    const oldIndex = row * oldColumns + column;

    if (row < oldRows && column < oldColumns) {
      return {
        ...createEmptyCell(),
        ...cells[oldIndex],
      };
    }

    return createEmptyCell();
  });
}

function alignToFlex(align: "left" | "center" | "right") {
  if (align === "center") {
    return "center";
  }

  if (align === "right") {
    return "flex-end";
  }

  return "flex-start";
}

function verticalAlignToFlex(align: "top" | "middle" | "bottom") {
  if (align === "middle") {
    return "center";
  }

  if (align === "bottom") {
    return "flex-end";
  }

  return "flex-start";
}

function normalizeDistribution(sizes: number[] | undefined, count: number, total: number) {
  const safeCount = Math.max(1, count);
  const fallback = total / safeCount;
  const normalized = Array.from({ length: safeCount }, (_, index) => sizes?.[index] ?? fallback);
  const currentTotal = sum(normalized);

  if (currentTotal <= 0) {
    return Array.from({ length: safeCount }, () => roundMm(fallback));
  }

  return normalized.map((size) => roundMm((size / currentTotal) * total));
}

function resizeDistribution(sizes: number[], count: number, total: number) {
  const safeCount = Math.max(1, count);
  if (safeCount === sizes.length) {
    return normalizeDistribution(sizes, safeCount, total);
  }

  if (safeCount < sizes.length) {
    return normalizeDistribution(sizes.slice(0, safeCount), safeCount, total);
  }

  const nextSizes = [...sizes];
  const fallback = total / safeCount;
  while (nextSizes.length < safeCount) {
    nextSizes.push(fallback);
  }

  return normalizeDistribution(nextSizes, safeCount, total);
}

function resizeTableSizes(sizes: number[], index: number, delta: number, minSize: number) {
  const nextSizes = [...sizes];
  const left = sizes[index];
  const right = sizes[index + 1];
  const availableDelta = Math.min(Math.max(delta, minSize - left), right - minSize);

  nextSizes[index] = roundMm(left + availableDelta);
  nextSizes[index + 1] = roundMm(right - availableDelta);

  return nextSizes;
}

function cumulativeSize(sizes: number[], endIndex: number) {
  return sum(sizes.slice(0, endIndex));
}

function sum(values: number[]) {
  return values.reduce((total, value) => total + value, 0);
}
