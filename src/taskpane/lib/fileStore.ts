import { openDB, IDBPDatabase, DBSchema } from "idb";
import { BlockIndex } from "./referenceDoc";

interface FileRecord {
  name: string;
  blocks: BlockIndex[];
}

interface DebateAddinDB extends DBSchema {
  files: {
    key: string;
    value: FileRecord;
  };
}

const DB_NAME = "debate-addin";
const DB_VERSION = 3; // bumped: blocks now include cachedOoxml; base64 removed
const STORE_NAME = "files";

export class FileStore {
  private db: IDBPDatabase<DebateAddinDB> | null = null;
  private loadedFiles: Map<string, BlockIndex[]> = new Map();

  async init(): Promise<void> {
    this.db = await openDB<DebateAddinDB>(DB_NAME, DB_VERSION, {
      upgrade(db, oldVersion) {
        // Clear any previous store — block format has changed across all versions
        if (db.objectStoreNames.contains(STORE_NAME)) {
          db.deleteObjectStore(STORE_NAME);
        }
        db.createObjectStore(STORE_NAME, { keyPath: "name" });
      },
    });

    const records = await this.db.getAll(STORE_NAME);
    for (const record of records) {
      this.loadedFiles.set(record.name, record.blocks);
    }
  }

  /**
   * Store pre-built blocks (from loadReferenceFile) for a file.
   * Parsing and OOXML extraction are done by the caller before this point.
   */
  async addFile(file: File, blocks: BlockIndex[]): Promise<void> {
    this.loadedFiles.set(file.name, blocks);
    if (this.db) {
      await this.db.put(STORE_NAME, { name: file.name, blocks });
    }
  }

  removeFile(name: string): void {
    this.loadedFiles.delete(name);
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

  getAllBlocks(enabledFiles?: string[]): BlockIndex[] {
    const files = enabledFiles ?? [...this.loadedFiles.keys()];
    return files.flatMap(name => this.loadedFiles.get(name) ?? []);
  }
}
