import {Component, Input, OnInit} from '@angular/core';
import {OleFile} from "../models/ole-file";
import {faTimes, faSkull, faCheck} from "@fortawesome/free-solid-svg-icons

@Component({
  selector: 'app-file-preview',
  templateUrl: './file-preview.component.html',
  styleUrls: ['./file-preview.component.scss']
})
export class FilePreviewComponent implements OnInit {
  @Input() file: OleFile | undefined;
  faTimes = faTimes;
  faSkull = faSkull;
  faCheck = faCheck;

  constructor() {
  }

  ngOnInit(): void {
  }

  getReadableFileSizeString(fileSizeInBytes: number): string {
    const byteUnits = [' KB', ' MB', ' GB', ' TB'];
    let i = -1;
    while (fileSizeInBytes > 1000) {
      fileSizeInBytes = fileSizeInBytes / 1000;
      i++;
    }

    return Math.max(fileSizeInBytes, 0.1).toFixed(1) + " " + byteUnits[i];
  }
}
