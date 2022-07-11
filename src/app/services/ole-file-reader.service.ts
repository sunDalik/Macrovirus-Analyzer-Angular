import {Injectable} from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class OleFileReaderService {
  constructor() {
  }

  private readByteArray(byteArray: Uint8Array, offset: OffsetValue, bytesToRead: number): Uint8Array {
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

class OffsetValue {
  value: number;

  constructor(val: number) {
    this.value = val;
  }
}
