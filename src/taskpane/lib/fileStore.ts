import { openDB, IDBPDatabase, DBSchema } from "idb";
import { BlockIndex } from "./referenceDoc";

interface FileRecord {
  name: string;
  blocks: BlockIndex[];
  stylesJson: string; // exportStylesFromJson output; "" if unavailable at load time
}

interface DebateAddinDB extends DBSchema {
  files: {
    key: string;
    value: FileRecord;
  };
}

const DB_NAME = "debate-addin";
const DB_VERSION = 4; // bumped: FileRecord now includes stylesJson
const STORE_NAME = "files";

export class FileStore {
  private db: IDBPDatabase<DebateAddinDB> | null = null;
  private loadedFiles: Map<string, BlockIndex[]> = new Map();
  private stylesJsonMap: Map<string, string> = new Map();

  async init(): Promise<void> {
    this.db = await openDB<DebateAddinDB>(DB_NAME, DB_VERSION, {
      upgrade(db) {
        // Clear any previous store — schema has changed across versions
        if (db.objectStoreNames.contains(STORE_NAME)) {
          db.deleteObjectStore(STORE_NAME);
        }
        db.createObjectStore(STORE_NAME, { keyPath: "name" });
      },
    });

    const records = await this.db.getAll(STORE_NAME);
    for (const record of records) {
      this.loadedFiles.set(record.name, record.blocks);
      this.stylesJsonMap.set(record.name, record.stylesJson ?? "");
    }
  }

  /**
   * Store pre-built blocks and style definitions for a file.
   * All heavy lifting (parsing, OOXML extraction, style export) is done
   * by the caller via loadReferenceFile() before this point.
   */
  async addFile(file: File, blocks: BlockIndex[], stylesJson: string): Promise<void> {
    this.loadedFiles.set(file.name, blocks);
    this.stylesJsonMap.set(file.name, stylesJson);
    if (this.db) {
      await this.db.put(STORE_NAME, { name: file.name, blocks, stylesJson });
    }
  }

  removeFile(name: string): void {
    this.loadedFiles.delete(name);
    this.stylesJsonMap.delete(name);
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

  /** Returns the cached exportStylesFromJson output, or "" if unavailable. */
  getStylesJson(name: string): string {
    return this.stylesJsonMap.get(name) ?? "";
  }
}
