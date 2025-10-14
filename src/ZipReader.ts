// SyncZipReader.ts
import * as fs from 'fs';
import * as zlib from 'zlib';

export class ZipConsts {
    // Local File Header
    static readonly LOCHDR = 30;
    static readonly LOCSIG = 0x04034b50;
    static readonly LOCVER = 4;
    static readonly LOCFLG = 6;
    static readonly LOCHOW = 8;
    static readonly LOCTIM = 10;
    static readonly LOCCRC = 14;
    static readonly LOCSIZ = 18;
    static readonly LOCLEN = 22;
    static readonly LOCNAM = 26;
    static readonly LOCEXT = 28;

    // Data Descriptor
    static readonly EXTSIG = 0x08074b50;
    static readonly EXTHDR = 16;
    static readonly EXTCRC = 4;
    static readonly EXTSIZ = 8;
    static readonly EXTLEN = 12;

    // Central Directory
    static readonly CENHDR = 46;
    static readonly CENSIG = 0x02014b50;
    static readonly CENVEM = 4;
    static readonly CENVER = 6;
    static readonly CENFLG = 8;
    static readonly CENHOW = 10;
    static readonly CENTIM = 12;
    static readonly CENCRC = 16;
    static readonly CENSIZ = 20;
    static readonly CENLEN = 24;
    static readonly CENNAM = 28;
    static readonly CENEXT = 30;
    static readonly CENCOM = 32;
    static readonly CENDSK = 34;
    static readonly CENATT = 36;
    static readonly CENATX = 38;
    static readonly CENOFF = 42;

    // End of Central Directory
    static readonly ENDHDR = 22;
    static readonly ENDSIG = 0x06054b50;
    static readonly ENDSIGFIRST = 0x50;
    static readonly ENDSUB = 8;
    static readonly ENDTOT = 10;
    static readonly ENDSIZ = 12;
    static readonly ENDOFF = 16;
    static readonly ENDCOM = 20;
    static readonly MAXFILECOMMENT = 0xffff;

    // ZIP64
    static readonly ENDL64HDR = 20;
    static readonly ENDL64SIG = 0x07064b50;
    static readonly ENDL64SIGFIRST = 0x50;
    static readonly ENDL64OFS = 8;

    static readonly END64HDR = 56;
    static readonly END64SIG = 0x06064b50;
    static readonly END64SIGFIRST = 0x50;
    static readonly END64SUB = 24;
    static readonly END64TOT = 32;
    static readonly END64SIZ = 40;
    static readonly END64OFF = 48;

    // Compression
    static readonly STORED = 0;
    static readonly DEFLATED = 8;

    // Flags
    static readonly FLG_ENTRY_ENC = 1;

    // Sentinel values
    static readonly EF_ZIP64_OR_32 = 0xffffffff;
    static readonly EF_ZIP64_OR_16 = 0xffff;

    // Extra field
    static readonly ID_ZIP64 = 0x0001;
}

export class ZipReader {
    private readonly buffer: Buffer;

    constructor(filePath: string) {
        const fd = fs.openSync(filePath, 'r');
        try {
            const size = fs.fstatSync(fd).size;
            this.buffer = Buffer.alloc(size);
            fs.readSync(fd, this.buffer, 0, size, 0);
        } finally {
            fs.closeSync(fd);
        }
    }

    listEntries(): string[] {
        const entries: string[] = [];
        const { offset, count } = this.findCentralDirectory();
        let pos = offset;
        for (let i = 0; i < count; i++) {
            if (pos + ZipConsts.CENHDR > this.buffer.length) break;
            if (this.buffer.readUInt32LE(pos) !== ZipConsts.CENSIG) break;

            const nameLen = this.buffer.readUInt16LE(pos + ZipConsts.CENNAM);
            const extraLen = this.buffer.readUInt16LE(pos + ZipConsts.CENEXT);
            const comLen = this.buffer.readUInt16LE(pos + ZipConsts.CENCOM);
            const name = this.buffer.subarray(pos + ZipConsts.CENHDR, pos + ZipConsts.CENHDR + nameLen).toString();
            entries.push(name);
            pos += ZipConsts.CENHDR + nameLen + extraLen + comLen;
        }
        return entries;
    }

    getEntry(name: string): Buffer | null {
        const { offset, count } = this.findCentralDirectory();
        let pos = offset;
        for (let i = 0; i < count; i++) {
            if (pos + ZipConsts.CENHDR > this.buffer.length) break;
            if (this.buffer.readUInt32LE(pos) !== ZipConsts.CENSIG) break;

            const nameLen = this.buffer.readUInt16LE(pos + ZipConsts.CENNAM);
            const extraLen = this.buffer.readUInt16LE(pos + ZipConsts.CENEXT);
            const comLen = this.buffer.readUInt16LE(pos + ZipConsts.CENCOM);
            const entryName = this.buffer.subarray(pos + ZipConsts.CENHDR, pos + ZipConsts.CENHDR + nameLen).toString();

            if (entryName === name) {
                const compressedSize = this.buffer.readUInt32LE(pos + ZipConsts.CENSIZ);
                const uncompressedSize = this.buffer.readUInt32LE(pos + ZipConsts.CENLEN);
                const localHeaderOffset = this.buffer.readUInt32LE(pos + ZipConsts.CENOFF);
                const method = this.buffer.readUInt16LE(pos + ZipConsts.CENHOW);

                // Read local header to get fname/extra lengths
                const fnameLen = this.buffer.readUInt16LE(localHeaderOffset + ZipConsts.LOCNAM);
                const fextraLen = this.buffer.readUInt16LE(localHeaderOffset + ZipConsts.LOCEXT);
                const dataStart = localHeaderOffset + ZipConsts.LOCHDR + fnameLen + fextraLen;
                let data = this.buffer.subarray(dataStart, dataStart + compressedSize);

                if (method === ZipConsts.DEFLATED) {
                    data = zlib.inflateRawSync(data);
                } else if (method !== ZipConsts.STORED) {
                    throw new Error(`Unsupported compression method: ${method}`);
                }

                if (data.length !== uncompressedSize) {
                    throw new Error(`Size mismatch in entry: ${name}`);
                }

                return data;
            }

            pos += ZipConsts.CENHDR + nameLen + extraLen + comLen;
        }
        return null;
    }

    private findCentralDirectory(): { offset: number; count: number } {
        const len = this.buffer.length;
        for (let i = len - ZipConsts.ENDHDR; i >= 0; i--) {
            if (this.buffer.readUInt32LE(i) === ZipConsts.ENDSIG) {
                const offset = this.buffer.readUInt32LE(i + ZipConsts.ENDOFF);
                const count = this.buffer.readUInt16LE(i + ZipConsts.ENDTOT);
                return { offset, count };
            }
        }
        throw new Error('Invalid ZIP: End of Central Directory not found');
    }
}
