import {Injectable} from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class FileReaderService {
  constructor() {
  }

  readInt(byteArray: Uint8Array, offset: OffsetValue, bytesToRead: number): number {
    const slicedArray = byteArray.slice(offset.value, offset.value + bytesToRead).buffer;
    let value;
    if (bytesToRead === 1) value = new DataView(slicedArray).getUint8(0);
    else if (bytesToRead === 2) value = new DataView(slicedArray).getUint16(0, true);
    else if (bytesToRead === 4) value = new DataView(slicedArray).getUint32(0, true);
    else if (bytesToRead > 4) value = Number(new DataView(slicedArray).getBigUint64(0, true));
    else throw Error("Invalid bytesToRead value: " + bytesToRead);
    offset.value += bytesToRead;
    return value;
  }

  skipStructure(buffer: Uint8Array, offset: OffsetValue, lengthBytes: number, elementSize = 1) {
    let length = this.readInt(buffer, offset, lengthBytes);
    let invalidLength = lengthBytes === 2 ? 0xFFFF : 0xFFFFFFFF;
    if (length !== invalidLength) {
      // @ts-ignore
      offset.value += length * elementSize;
    }
  }

  byteArrayToStr(array: Uint8Array): string {
    // utf16
    let string = "";
    for (let i = 0; i < array.length; i++) {
      const code = array[i];
      if (code === 0) continue;
      string += String.fromCharCode(code);
    }
    return string;
  }

  readByteArray(byteArray: Uint8Array, offset: OffsetValue, bytesToRead: number): Uint8Array {
    const slicedArray = byteArray.slice(offset.value, offset.value + bytesToRead);
    offset.value += bytesToRead;
    return slicedArray;
  }

  isOleFile(contents: Uint8Array): boolean {
    const magic = this.readByteArray(contents, new OffsetValue(0), 8);
    return magic[0] === 0xD0 && magic[1] === 0xCF && magic[2] === 0x11 && magic[3] === 0xE0
      && magic[4] === 0xA1 && magic[5] === 0xB1 && magic[6] === 0x1A && magic[7] === 0xE1;
  }

  isZipFile(contents: Uint8Array): boolean {
    const magic = this.readByteArray(contents, new OffsetValue(0), 4);
    return magic[0] === 0x50 && magic[1] === 0x4B && magic[2] === 0x03 && magic[3] === 0x04;
  }
}

export class OffsetValue {
  value: number;

  constructor(val: number) {
    this.value = val;
  }
}
