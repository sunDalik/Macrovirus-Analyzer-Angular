import {Component, Input, OnInit, Output, EventEmitter} from '@angular/core';
import {OleFile} from "../models/ole-file";
import {faTimes, faSkull, faCheck, faUpload} from "@fortawesome/free-solid-svg-icons"
import {Router} from "@angular/router";

@Component({
  selector: 'app-file-scroller',
  templateUrl: './file-scroller.component.html',
  styleUrls: ['./file-scroller.component.scss']
})
export class FileScrollerComponent implements OnInit {
  @Input() files: Array<OleFile> = [];
  @Output() fileDisplayEvent = new EventEmitter<OleFile>();
  @Output() newFilesUploaded = new EventEmitter<Array<OleFile>>();
  faTimes = faTimes;
  faSkull = faSkull;
  faCheck = faCheck;
  faUpload = faUpload;

  constructor(public router: Router) {
  }

  ngOnInit(): void {
  }

  onFilesUpload(files: Array<OleFile>) {
    this.newFilesUploaded.emit(files);
    this.files = files;
  }
}
