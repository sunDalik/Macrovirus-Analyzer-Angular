import {FileReaderService, OffsetValue} from "../services/file-reader.service";
import {FuncType, VbaAnalysisService} from "../services/vba-analysis.service";

export class MacroModule {
  name: string = "";
  sourceCode: string = "";
  pcode: string = "";
}

enum DirEntryType {
  UNKNOWN = 0,
  STORAGE = 1,
  STREAM = 2,
  ROOT = 5
}

class DirEntry {
  id: number;
  name = "";
  type = DirEntryType.UNKNOWN;
  children: Array<DirEntry> = [];
  startingSector: number = 0xFFFFFFFF;
  streamSize: number = 0;

  //temporary fields needed to build a proper tree later
  leftSiblingId: number | null = null;
  rightSiblingId: number | null = null;
  childId: number | null = null;

  constructor(id: number) {
    this.id = id;
  }
}

export class VBAFunction {
  dependencies: Array<number> = [];
  body: Array<string> = [];

  constructor(public id: number, public name: string, public type: FuncType) {
  }
}

export class OleFile {
  readError = false;
  isMalicious = false;
  macroModules: Array<MacroModule> = [];
  analysisResult = "";
  VBAFunctions: Array<VBAFunction> = []

  private sectorSize = 512;
  private miniSectorSize = 64;
  private miniStreamCutoff = 4096;
  private DIFAT = new Uint8Array(0);
  private FAT = new Uint8Array(0);
  private miniFAT = new Uint8Array(0);
  private miniStream = new Uint8Array(0);
  private fileTree: DirEntry | undefined;

  constructor(public file: File, public binaryContent: Uint8Array, public fileReaderService: FileReaderService, public analysisService: VbaAnalysisService) {
    try {
      this.processFile();
    } catch (e) {
      console.log("ERROR READING OLEFILE " + this.file.name);
      this.readError = true;
    }
    this.analysisResult = this.analysisService.analyzeFile(this);
  }

  processFile() {
    const binContent = this.binaryContent;
    // https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
    const offset = {value: 0};
    const headerSignature = this.fileReaderService.readInt(binContent, offset, 8); // Must be 0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1
    const headerCLSID = this.fileReaderService.readByteArray(binContent, offset, 16);
    const minorVersion = this.fileReaderService.readInt(binContent, offset, 2);
    const majorVersion = this.fileReaderService.readInt(binContent, offset, 2);
    const byteOrder = this.fileReaderService.readInt(binContent, offset, 2);
    const sectorShift = this.fileReaderService.readInt(binContent, offset, 2);
    this.sectorSize = 2 ** sectorShift;
    const miniSectorShift = this.fileReaderService.readInt(binContent, offset, 2);
    this.miniSectorSize = 2 ** miniSectorShift;
    const reserved = this.fileReaderService.readByteArray(binContent, offset, 6);
    const numberOfDirectorySectors = this.fileReaderService.readInt(binContent, offset, 4);
    const numberOfFATSectors = this.fileReaderService.readInt(binContent, offset, 4);
    const firstDirectorySector = this.fileReaderService.readInt(binContent, offset, 4);
    const transactionSignatureNumber = this.fileReaderService.readInt(binContent, offset, 4);
    this.miniStreamCutoff = this.fileReaderService.readInt(binContent, offset, 4);
    const firstMiniFATSector = this.fileReaderService.readInt(binContent, offset, 4);
    const numberOfMiniFATSectors = this.fileReaderService.readInt(binContent, offset, 4);
    const firstDIFATSector = this.fileReaderService.readInt(binContent, offset, 4);
    const numberOfDIFATSectors = this.fileReaderService.readInt(binContent, offset, 4);

    this.DIFAT = this.readDIFAT(numberOfDIFATSectors, firstDIFATSector);
    this.FAT = this.readFAT(numberOfFATSectors);
    this.miniFAT = this.readSectorChainFAT(firstMiniFATSector);
    this.fileTree = this.readFileTree(firstDirectorySector);
    //@ts-ignore
    this.miniStream = this.readSectorChainFAT(this.fileTree.startingSector);

    const vbaFolder = this.findByPath(this.fileTree, "MACROS/VBA")
      || this.findByPath(this.fileTree, "VBA")
      || this.findByPath(this.fileTree, "_VBA_PROJECT_CUR/VBA");

    //const wb = this.findByPath(this.fileTree, "Workbook");
    //console.log(this.fileReaderService.byteArrayToStr(this.readStream(wb.startingSector, wb.streamSize)));

    if (vbaFolder) {
      const modules = [];
      let dirStream = null;
      let vbaProjectStream = null;
      for (const child of vbaFolder.children) {
        if (!["_VBA_PROJECT", "DIR"].includes(child.name.toUpperCase()) && !child.name.toUpperCase().startsWith("__SRP_")) {
          modules.push(child);
        }
        if (child.name.toUpperCase() === "DIR") {
          dirStream = this.readStream(child.startingSector, child.streamSize);
        }
        if (child.name.toUpperCase() === "_VBA_PROJECT") {
          vbaProjectStream = this.readStream(child.startingSector, child.streamSize);
        }
      }

      if (dirStream === null) throw Error("Dir stream is null");
      dirStream = this.decompressVBASourceCode(dirStream);
      const dirOffset = {value: 0};

      //read project information record
      dirOffset.value += 38;

      dirOffset.value += 2;
      const projectNameRecordSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
      dirOffset.value += projectNameRecordSize;

      dirOffset.value += 2;
      const docStringSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
      dirOffset.value += docStringSize;
      dirOffset.value += 2;
      const docStringUnicodeSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
      dirOffset.value += docStringUnicodeSize;

      dirOffset.value += 2;
      const helpFileSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
      dirOffset.value += helpFileSize;
      dirOffset.value += 2;
      const helpFileSize2 = this.fileReaderService.readInt(dirStream, dirOffset, 4);
      dirOffset.value += helpFileSize2;

      dirOffset.value += 32;

      // Optional constants field
      //@ts-ignore
      const constantsId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
      if (constantsId === 0x000C) {
        const constantsSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        dirOffset.value += constantsSize;
        dirOffset.value += 2;

        const constantsSizeUnicode = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        dirOffset.value += constantsSizeUnicode;
      } else {
        dirOffset.value -= 2;
      }

      //read project references record
      let attempt = 0;
      while (attempt++ < 999999) {
        const nameReadResult = this.readNameReference(dirOffset, dirStream);
        if (nameReadResult === false) break;


        const referenceRecordId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
        if (referenceRecordId === 0x002F) {
          this.readControlReference(dirOffset, dirStream);
        } else if (referenceRecordId === 0x0033) {
          this.readOriginalReference(dirOffset, dirStream);
        } else if (referenceRecordId === 0x000D) {
          this.readRegisteredReference(dirOffset, dirStream);
        } else if (referenceRecordId === 0x000E) {
          this.readProjectReference(dirOffset, dirStream);
        }
      }

      // reading modules records
      // we've already read the id
      const countSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
      const count = this.fileReaderService.readInt(dirStream, dirOffset, countSize);
      dirOffset.value += 8;
      const moduleRecordsArray = [];
      for (let i = 0; i < count; i++) {
        //module name
        const moduleNameId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
        const moduleNameSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        dirOffset.value += moduleNameSize;

        //module name unicode OPTIONAL
        const moduleNameUnicodeId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
        if (moduleNameUnicodeId === 0x0047) {
          const moduleNameUnicodeSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
          dirOffset.value += moduleNameUnicodeSize;
        } else {
          dirOffset.value -= 2;
        }

        //module stream name
        dirOffset.value += 2;
        const streamNameSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        const streamName = this.fileReaderService.readByteArray(dirStream, dirOffset, streamNameSize);
        dirOffset.value += 2;
        const streamNameUnicodeSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        const streamNameUnicode = this.fileReaderService.readByteArray(dirStream, dirOffset, streamNameUnicodeSize);

        //doc string
        dirOffset.value += 2;
        const docStringSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        dirOffset.value += docStringSize;
        dirOffset.value += 2;
        const docStringUnicodeSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        dirOffset.value += docStringUnicodeSize;

        //module offset
        const sourceOffsetId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
        const sourceOffsetSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
        if (sourceOffsetSize !== 4) console.log("SOURCE OFFSET SIZE != 4");
        const sourceOffset = this.fileReaderService.readInt(dirStream, dirOffset, 4);

        dirOffset.value += 10; // help context
        dirOffset.value += 8; // module cookie

        dirOffset.value += 6; // module type

        // optional read only field
        const readOnlyId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
        if (readOnlyId === 0x0025) {
          dirOffset.value += 4;
        } else {
          dirOffset.value -= 2;
        }

        // optional module private field
        const privateId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
        if (privateId === 0x0028) {
          dirOffset.value += 4;
        } else {
          dirOffset.value -= 2;
        }

        dirOffset.value += 2; // terminator
        dirOffset.value += 4; // reserved

        //todo secondModule seems to be wrong? empty name and ridiculous offset
        moduleRecordsArray.push({
          name: this.fileReaderService.byteArrayToStr(streamName),
          sourceOffset: sourceOffset,
          nameUnicode: this.fileReaderService.byteArrayToStr(streamNameUnicode)
        });
      }

      this.macroModules = [];

      for (const module of modules) {
        const macroModule = new MacroModule();
        const dataArray = this.readStream(module.startingSector, module.streamSize);
        const moduleRecord = moduleRecordsArray.find(m => m.name.toUpperCase() === module.name.toUpperCase());
        if (!moduleRecord) continue;
        macroModule.name = moduleRecord.name;
        macroModule.sourceCode = this.fileReaderService.byteArrayToStr(this.decompressVBASourceCode(dataArray.slice(moduleRecord.sourceOffset)));
        macroModule.sourceCode = this.removeAttributes(macroModule.sourceCode);
        try {
          //macroModule.pcode = disassemblePCode(dataArray, vbaProjectStream, dirStream);
        } catch (e) {
          console.log("PCODE DISASSEMBLING ERROR");
          //macroModule.pcode = [];
        }
        this.macroModules.push(macroModule);
      }
    }
  }

  getFormat(): string {
    return this.file.name.split(".")[1];
  }


  findByPath(tree: DirEntry, path: string): DirEntry | undefined {
    const pathArray = path.split("/");
    let currentFile: DirEntry | undefined = tree;
    for (const segment of pathArray) {
      currentFile = currentFile.children.find(f => f.name.toUpperCase() === segment.toUpperCase());
      if (!currentFile) break;
    }
    return currentFile;
  }

  readNameReference(dirOffset: OffsetValue, dirStream: Uint8Array): boolean {
    const id = this.fileReaderService.readInt(dirStream, dirOffset, 2);
    if (id === 0x000F) return false; // value 0x000F indicates end of references array and start of module records
    const nameSize = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += nameSize;
    dirOffset.value += 2;
    const nameSizeUnicode = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += nameSizeUnicode;
    return true;
  }

  readControlReference(dirOffset: OffsetValue, dirStream: Uint8Array) {
    const sizeTwiddled = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    const sizeLibidTwiddled = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += sizeLibidTwiddled;
    dirOffset.value += 6;
    const reserved = this.fileReaderService.readInt(dirStream, dirOffset, 2);
    // Optional Name Record
    if (reserved === 0x0016) {
      dirOffset.value -= 2;
      this.readNameReference(dirOffset, dirStream);
      dirOffset.value += 2;
    }
    const sizeExtended = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    const sizeLibidExtended = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += sizeLibidExtended;
    dirOffset.value += 26;
  }

  readOriginalReference(dirOffset: OffsetValue, dirStream: Uint8Array) {
    const sizeLibidOriginal = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += sizeLibidOriginal;

    // OriginalReference might be the beginning of a ControlReference
    const referenceRecordId = this.fileReaderService.readInt(dirStream, dirOffset, 2);
    if (referenceRecordId === 0x2F) {
      this.readControlReference(dirOffset, dirStream);
    }
  }

  readRegisteredReference(dirOffset: OffsetValue, dirStream: Uint8Array) {
    const size = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    const sizeLibid = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += sizeLibid;
    dirOffset.value += 6;
  }

  readProjectReference(dirOffset: OffsetValue, dirStream: Uint8Array) {
    const size = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    const sizeLibidAbsolute = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += sizeLibidAbsolute;
    const sizeLibidRelative = this.fileReaderService.readInt(dirStream, dirOffset, 4);
    dirOffset.value += sizeLibidRelative;
    dirOffset.value += 6;
  }

  readStream(sectorNumber: number, streamSize: number): Uint8Array {
    if (streamSize >= this.miniStreamCutoff) {
      return this.readSectorChainFAT(sectorNumber, streamSize);
    } else {
      return this.readMiniSectorChain(sectorNumber, streamSize);
    }
  }

  readSectorFAT(sectorNumber: number): Uint8Array {
    const offset = (sectorNumber + 1) * this.sectorSize;
    return this.binaryContent.slice(offset, offset + this.sectorSize);
  }

  readSectorChainFAT(startSectorNumber: number, streamSize = -1): Uint8Array {
    if (startSectorNumber === 0xFFFFFFFE || startSectorNumber === 0xFFFFFFFF) return new Uint8Array(0);
    const sectorIndexesArray = [startSectorNumber];
    let attempt = 0;
    while (attempt++ < 999999) {
      const nextSector = this.fileReaderService.readInt(this.FAT, {value: sectorIndexesArray[sectorIndexesArray.length - 1] * 4}, 4);
      if (nextSector === 0xFFFFFFFE || nextSector === 0xFFFFFFFF) break;
      sectorIndexesArray.push(nextSector);
    }

    let resultArrayLength = streamSize === -1 ? sectorIndexesArray.length * this.sectorSize : streamSize;
    resultArrayLength = Number(resultArrayLength);
    const resultArray = new Uint8Array(resultArrayLength);
    let bytesRead = 0;
    for (let i = 0; i < sectorIndexesArray.length; i++) {
      const sectorNumber = sectorIndexesArray[i];
      const offset = (sectorNumber + 1) * this.sectorSize;
      for (let b = 0; b < this.sectorSize; b++) {
        resultArray[i * this.sectorSize + b] = this.binaryContent[offset + b];
        bytesRead++;
        if (bytesRead >= resultArrayLength) return resultArray;
      }
    }
    return resultArray;
  }

  readMiniSector(sectorNumber: number): Uint8Array {
    const offset = sectorNumber * this.miniSectorSize;
    return this.miniStream.slice(offset, offset + this.miniSectorSize);
  }

  readMiniSectorChain(startSectorNumber: number, streamSize = -1): Uint8Array {
    const sectorIndexesArray = [startSectorNumber];
    let attempt = 0;
    while (attempt++ < 9999999) {
      const nextSector = this.fileReaderService.readInt(this.miniFAT, {value: sectorIndexesArray[sectorIndexesArray.length - 1] * 4}, 4);
      if (nextSector === 0xFFFFFFFE) break;
      sectorIndexesArray.push(nextSector);
    }

    let resultArrayLength = streamSize === -1 ? sectorIndexesArray.length * this.miniSectorSize : streamSize;
    resultArrayLength = Number(resultArrayLength);
    const resultArray = new Uint8Array(resultArrayLength);
    let bytesRead = 0;
    for (let i = 0; i < sectorIndexesArray.length; i++) {
      const sectorNumber = sectorIndexesArray[i];
      const offset = sectorNumber * this.miniSectorSize;
      for (let b = 0; b < this.miniSectorSize; b++) {
        resultArray[i * this.miniSectorSize + b] = this.miniStream[offset + b];
        bytesRead++;
        if (bytesRead >= resultArrayLength) return resultArray;
      }
    }
    return resultArray;
  }

  readDIFAT(numberOfDIFATSectors: number, firstDIFATSector: number): Uint8Array {
    const headerDIFATSize = 109 * 4;
    const DIFAT = new Uint8Array(headerDIFATSize + numberOfDIFATSectors * (this.sectorSize - 4));

    const headerDIFAT = this.fileReaderService.readByteArray(this.binaryContent, {value: 76}, headerDIFATSize);
    for (let b = 0; b < headerDIFAT.length; b++) {
      DIFAT[b] = headerDIFAT[b];
    }

    let DIFATSector = firstDIFATSector;
    for (let i = 0; i < numberOfDIFATSectors; i++) {
      const sector = this.readSectorFAT(DIFATSector);
      for (let b = 0; b < sector.length - 4; b++) {
        DIFAT[headerDIFATSize + i * (this.sectorSize - 4) + b] = sector[b];
      }
      DIFATSector = this.fileReaderService.readInt(sector, {value: this.sectorSize - 4}, 4);
    }

    return DIFAT;
  }


  readFAT(numberOfFATSectors: number): Uint8Array {
    const FAT = new Uint8Array(numberOfFATSectors * this.sectorSize);
    let offset = {value: 0};
    for (let i = 0; i < numberOfFATSectors; i++) {
      const sectorNumber = this.fileReaderService.readInt(this.DIFAT, offset, 4);
      const sector = this.readSectorFAT(sectorNumber);
      for (let b = 0; b < sector.length; b++) {
        FAT[i * this.sectorSize + b] = sector[b];
      }
    }
    return FAT;
  }


  readFileTree(firstDirectorySector: number): DirEntry {
    const directoryEntrySize = 128;
    const directoriesData = this.readSectorChainFAT(firstDirectorySector);

    const directoryEntries = [];
    for (let i = 0; i < directoriesData.length; i += directoryEntrySize) {
      const dirEntry = new DirEntry(i / directoryEntrySize);
      const offset = {value: i + 64};
      const nameFieldLength = this.fileReaderService.readInt(directoriesData, offset, 2);

      const nameLength = Math.max(nameFieldLength / 2 - 1, 0);
      const nameArray = new Uint8Array(nameLength);
      for (let j = 0; j < nameLength; j++) {
        nameArray[j] = directoriesData[i + j * 2];
      }
      dirEntry.name = this.fileReaderService.byteArrayToStr(nameArray);

      const objectType = this.fileReaderService.readInt(directoriesData, offset, 1);
      if (objectType === 0) dirEntry.type = DirEntryType.UNKNOWN;
      else if (objectType === 1) dirEntry.type = DirEntryType.STORAGE;
      else if (objectType === 2) dirEntry.type = DirEntryType.STREAM;
      else if (objectType === 5) dirEntry.type = DirEntryType.ROOT;

      const colorFlag = this.fileReaderService.readInt(directoriesData, offset, 1);
      dirEntry.leftSiblingId = this.fileReaderService.readInt(directoriesData, offset, 4);
      dirEntry.rightSiblingId = this.fileReaderService.readInt(directoriesData, offset, 4);
      dirEntry.childId = this.fileReaderService.readInt(directoriesData, offset, 4);
      const CLSID = this.fileReaderService.readByteArray(directoriesData, offset, 16);
      const stateBits = this.fileReaderService.readByteArray(directoriesData, offset, 4);
      const creationTime = this.fileReaderService.readInt(directoriesData, offset, 8);
      const modifiedTime = this.fileReaderService.readInt(directoriesData, offset, 8);
      dirEntry.startingSector = this.fileReaderService.readInt(directoriesData, offset, 4);
      dirEntry.streamSize = this.fileReaderService.readInt(directoriesData, offset, 8);
      if (dirEntry.type !== DirEntryType.UNKNOWN) {
        directoryEntries.push(dirEntry);
      }
    }

    const findSiblings = (dir: DirEntry | undefined, entries: Array<DirEntry>) => {
      if (dir === undefined) return [];
      let siblings = [dir];
      if (dir.leftSiblingId !== null && dir.leftSiblingId !== 0xFFFFFFFF) {
        siblings = siblings.concat(findSiblings(entries.find(d => d.id === dir.leftSiblingId), entries));
      }
      if (dir.rightSiblingId !== null && dir.rightSiblingId !== 0xFFFFFFFF) {
        siblings = siblings.concat(findSiblings(entries.find(d => d.id === dir.rightSiblingId), entries));
      }
      return siblings;
    };

    const findChildren = (dir: DirEntry, entries: Array<DirEntry>) => {
      if (dir.childId !== null && dir.childId !== 0xFFFFFFFF) {
        dir.children = findSiblings(entries.find(d => d.id === dir.childId), entries);
      }

      for (const child of dir.children) {
        findChildren(child, entries);
      }
    };

    findChildren(directoryEntries[0], directoryEntries);
    return directoryEntries[0];
  }

  decompressVBASourceCode(byteCompressedArray: Uint8Array): Uint8Array {
    const decompressedArray: Array<number> = [];
    let compressedCurrent = 0;

    const sigByte = byteCompressedArray[compressedCurrent];
    if (sigByte !== 0x01) {
      throw new Error("Invalid signature byte");
    }

    compressedCurrent++;

    while (compressedCurrent < byteCompressedArray.length) {
      let compressedChunkStart = compressedCurrent;

      // first 16 bits
      // the header is in LE and the rightmost bit is the 0th.. makes sense right?
      let compressedChunkHeader = this.fileReaderService.readInt(byteCompressedArray, {value: compressedChunkStart}, 2);

      // first 12 bits
      let chunkSize = (compressedChunkHeader & 0x0FFF) + 3;

      const chunkSignature = (compressedChunkHeader >> 12) & 0x07;
      if (chunkSignature !== 0b011) {
        return new Uint8Array(decompressedArray); //?????
        //throw new Error("Invalid CompressedChunkSignature in VBA compressed stream");
      }

      // 15th flag
      let chunkFlag = (compressedChunkHeader >> 15) & 0x01;
      const compressedEnd = Math.min(byteCompressedArray.length, compressedChunkStart + chunkSize);
      compressedCurrent = compressedChunkStart + 2;
      const isCompressed = chunkFlag === 1;
      if (!isCompressed) {
        for (let i = 0; i < 4096; i++) {
          decompressedArray.push(byteCompressedArray[compressedCurrent]);
          compressedCurrent++;
        }
      } else {
        //???
        const decompressedChunkStart = decompressedArray.length;
        while (compressedCurrent < compressedEnd) {
          const flagByte = byteCompressedArray[compressedCurrent];
          compressedCurrent++;
          for (let bitIndex = 0; bitIndex < 8; bitIndex++) {
            if (compressedCurrent >= compressedEnd) break;
            const flagBit = (flagByte >> bitIndex) & 1;
            const isLiteral = flagBit === 0;
            if (isLiteral) {
              decompressedArray.push(byteCompressedArray[compressedCurrent]);
              compressedCurrent++;
            } else {
              const copyToken = this.fileReaderService.readInt(byteCompressedArray, {value: compressedCurrent}, 2);
              const decompressedCurrent = decompressedArray.length;
              const difference = decompressedCurrent - decompressedChunkStart;
              const bitCount = Math.max(Math.ceil(Math.log2(difference)), 4);
              const lengthMask = 0xFFFF >> bitCount;
              const offsetMask = ~lengthMask;
              const length = (copyToken & lengthMask) + 3;
              const tmp1 = copyToken & offsetMask;
              const tmp2 = 16 - bitCount;
              const offset = (tmp1 >> tmp2) + 1;
              const copySource = decompressedArray.length - offset;
              for (let i = copySource; i < copySource + length; i++) {
                decompressedArray.push(decompressedArray[i]);
              }
              compressedCurrent += 2;
            }
          }
        }
      }
    }

    return new Uint8Array(decompressedArray);
  }

  removeAttributes(code: string): string {
    const codeArray = code.split("\n");
    for (let i = 0; i < codeArray.length; i++) {
      if (codeArray[i].startsWith("Attribute VB_")) {
        codeArray[i] = "";
      } else {
        codeArray[i] += "\n";
      }
    }
    return codeArray.join("");
  }
}
