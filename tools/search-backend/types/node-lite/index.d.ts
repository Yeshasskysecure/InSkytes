declare const Buffer: {
  from(value: ArrayBuffer | string, encoding?: string): Buffer;
};

interface Buffer extends Uint8Array {
  byteLength: number;
  toString(encoding?: string): string;
}

declare const require: {
  (id: string): any;
};

declare const process: {
  argv: string[];
  exit(code?: number): never;
};

declare module 'fs' {
  const fs: any;
  export = fs;
}

declare module 'path' {
  const path: any;
  export = path;
}

declare module 'mammoth' {
  const mammoth: any;
  export = mammoth;
}

declare module 'xlsx' {
  const XLSX: any;
  export = XLSX;
}
