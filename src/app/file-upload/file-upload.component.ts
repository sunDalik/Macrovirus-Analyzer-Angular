import {Component, OnInit} from '@angular/core';
import {Router} from "@angular/router";
import {OleFileReaderService} from "../services/ole-file-reader.service";
import * as JSZip from "jszip";
import {OleFile} from "../models/ole-file";

@Component({
  selector: 'app-file-upload',
  templateUrl: './file-upload.component.html',
  styleUrls: ['./file-upload.component.scss']
})
export class FileUploadComponent implements OnInit {

  constructor(private oleFileReaderService: OleFileReaderService,
              private router: Router) {
  }

  ngOnInit(): void {
  }

  uploadFiles(e: Event) {
    const fileList = (<HTMLInputElement>e?.target)?.files ?? new FileList();
    if (fileList.length === 0) return;

    let pendingFiles = 0;
    const processedFiles: Array<OleFile | null> = [];

    const tryFinish = (newFile: OleFile | null) => {
      processedFiles.push(newFile);
      if (processedFiles.length >= fileList.length) {
        this.router.navigate(['result'], {state: {files: processedFiles.filter(f => f !== null)}});
      }
    }

    for (let i = 0; i < fileList.length; i++) {
      const file: File = fileList[i];
      if (file === null) {
        tryFinish(null);
        continue;
      }

      if (file.name.split(".")[1] === "zip") {
        //TODO Analyze all ole files inside the archive if file extension is .zip (or .rar or whatever?)
        //TODO Automatically try "infected" password if archive has password
      }

      const reader = new FileReader();
      reader.addEventListener('load', e => {
        pendingFiles--;
        const result = e?.target?.result;
        if (result == undefined) {
          tryFinish(null);
          return;
        }
        const contents = new Uint8Array(<ArrayBuffer>result);
        if (this.oleFileReaderService.isOleFile(contents)) {
          tryFinish(new OleFile(file, contents));
        } else if (this.oleFileReaderService.isZipFile(contents)) {
          const jsZip = new JSZip();
          jsZip.loadAsync(file)
            .then(zip => {
              for (const zipFileName in zip.files) {
                //TODO vbaProject.bin can potentially be renamed so you need to check for this
                if (zipFileName.endsWith("vbaProject.bin")) {
                  const currentFile = zip.files[zipFileName];
                  return currentFile.async("array");
                }
              }
              return [];
            }, () => {
              console.log("Not a valid zip file");
              //file.oleFile = null;
              tryFinish(null);
            })
            .then(content => {
              tryFinish(new OleFile(file, new Uint8Array(<number[]>content)));
            });
        }
      });
      reader.readAsArrayBuffer(file);
      pendingFiles++;
    }
  }
}
