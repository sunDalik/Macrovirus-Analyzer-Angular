export class OleFile {
  file: File;
  binaryContent: Uint8Array;
  public readError = false;
  public isMalicious = false;

  constructor(file: File, binaryContent: Uint8Array) {
    this.file = file;
    this.binaryContent = binaryContent;
    this.processFile();
  }

  processFile() {

  }
}
