import { openDB, IDBPDatabase, DBSchema } from "idb";
import { DebateBlock, parseDebateFile } from "./parser";

interface FileRecord {
  name: string;
  blocks: DebateBlock[];
}

interface DebateAddinDB extends DBSchema {
  files: {
    key: string;
    value: FileRecord;
  };
}

const DB_NAME = "debate-addin";
const DB_VERSION = 1;
const STORE_NAME = "files";

export class FileStore {
  private db: IDBPDatabase<DebateAddinDB> | null = null;
  private loadedFiles: Map<string, DebateBlock[]> = new Map();

  async init(): Promise<void> {
    this.db = await openDB<DebateAddinDB>(DB_NAME, DB_VERSION, {
      upgrade(db) {
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME, { keyPath: "name" });
        }
      },
    });

    // Restore all previously persisted files into in-memory map
    const records = await this.db.getAll(STORE_NAME);
    for (const record of records) {
      this.loadedFiles.set(record.name, record.blocks);
    }
  }

  async addFile(file: File): Promise<void> {
    const blocks = await parseDebateFile(file);
    this.loadedFiles.set(file.name, blocks);
    if (this.db) {
      await this.db.put(STORE_NAME, { name: file.name, blocks });
    }
  }

  removeFile(name: string): void {
    this.loadedFiles.delete(name);
    if (this.db) {
      // Fire-and-forget — deletion failures are non-critical
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
}
