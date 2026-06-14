import type { DocumentTemplate } from "./types";

const DB_NAME = "ldd-document-designer";
const STORE_NAME = "templates";
const DRAFT_STORE_NAME = "drafts";
const DB_VERSION = 2;

export type DocumentDraft = {
  id: string;
  savedAt: string;
  template: DocumentTemplate;
};

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(DRAFT_STORE_NAME)) {
        db.createObjectStore(DRAFT_STORE_NAME, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

export async function loadDraft(templateId: string): Promise<DocumentDraft | null> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DRAFT_STORE_NAME, "readonly");
    const store = transaction.objectStore(DRAFT_STORE_NAME);
    const request = store.get(templateId);

    request.onsuccess = () => resolve((request.result as DocumentDraft | undefined) ?? null);
    request.onerror = () => reject(request.error);
  });
}

export async function loadDrafts(): Promise<DocumentDraft[]> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DRAFT_STORE_NAME, "readonly");
    const store = transaction.objectStore(DRAFT_STORE_NAME);
    const request = store.getAll();

    request.onsuccess = () => resolve(request.result as DocumentDraft[]);
    request.onerror = () => reject(request.error);
  });
}

export async function saveDraft(template: DocumentTemplate): Promise<void> {
  const db = await openDb();
  const draft: DocumentDraft = {
    id: template.id,
    savedAt: new Date().toISOString(),
    template,
  };

  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DRAFT_STORE_NAME, "readwrite");
    transaction.objectStore(DRAFT_STORE_NAME).put(draft);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

export async function clearDraft(templateId: string): Promise<void> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DRAFT_STORE_NAME, "readwrite");
    transaction.objectStore(DRAFT_STORE_NAME).delete(templateId);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

export async function loadTemplates(): Promise<DocumentTemplate[]> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readonly");
    const store = transaction.objectStore(STORE_NAME);
    const request = store.getAll();

    request.onsuccess = () => {
      const templates = request.result as DocumentTemplate[];
      resolve(templates.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt)));
    };
    request.onerror = () => reject(request.error);
  });
}

export async function saveTemplate(template: DocumentTemplate): Promise<void> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readwrite");
    transaction.objectStore(STORE_NAME).put(template);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}

export async function deleteTemplate(id: string): Promise<void> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, "readwrite");
    transaction.objectStore(STORE_NAME).delete(id);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
  });
}
