import { openDB, IDBPDatabase, DBSchema } from "idb";
import { DebateBlock, parseDebateFile } from "./parser";

interface FileRecord {
  name: string;
  blocks: DebateBlock[];
  base64: string;
}

interface DebateAddinDB extends DBSchema {
  files: {
    key: string;
    value: FileRecord;
  };
}

const DB_NAME = "debate-addin";
const DB_VERSION = 2; // bumped: FileRecord now includes base64
const STORE_NAME = "files";

export class FileStore {
  private db: IDBPDatabase<DebateAddinDB> | null = null;
  private loadedFiles: Map<string, DebateBlock[]> = new Map();
  private base64Map: Map<string, string> = new Map();

  async init(): Promise<void> {
    this.db = await openDB<DebateAddinDB>(DB_NAME, DB_VERSION, {
      upgrade(db, oldVersion) {
        // Clear the old store when upgrading from v1 — the block format changed
        // (rawOoxml removed, paragraphStart/End added) so old records are unusable.
        if (oldVersion < 2 && db.objectStoreNames.contains(STORE_NAME)) {
          db.deleteObjectStore(STORE_NAME);
        }
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME, { keyPath: "name" });
        }
      },
    });

    const records = await this.db.getAll(STORE_NAME);
    for (const record of records) {
      this.loadedFiles.set(record.name, record.blocks);
      this.base64Map.set(record.name, record.base64);
    }
  }

  async addFile(file: File): Promise<void> {
    // Convert file to base64 for later paste operations
    const arrayBuffer = await file.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    let binary = "";
    for (let i = 0; i < bytes.length; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    const base64 = btoa(binary);

    const blocks = await parseDebateFile(file);
    this.loadedFiles.set(file.name, blocks);
    this.base64Map.set(file.name, base64);

    if (this.db) {
      await this.db.put(STORE_NAME, { name: file.name, blocks, base64 });
    }
  }

  removeFile(name: string): void {
    this.loadedFiles.delete(name);
    this.base64Map.delete(name);
    if (this.db) {
      this.db.delete(STORE_NAME, name).catch(console.error);
    }
  }

  getFileNames(): string[] {
    return [...this.loadedFiles.keys()];
  }

  getBlockCount(name: string): number {
    return this.loadedFiles.get(name)?.length ?? 0;
  }

  getAllBlocks(enabledFiles?: string[]): DebateBlock[] {
    const files = enabledFiles ?? [...this.loadedFiles.keys()];
    return files.flatMap(name => this.loadedFiles.get(name) ?? []);
  }

  getBase64(name: string): string | null {
    return this.base64Map.get(name) ?? null;
  }
}
