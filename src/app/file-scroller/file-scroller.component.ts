import {Component, Input, OnInit, Output, EventEmitter} from '@angular/core';
import {OleFile} from "../models/ole-file";
import {faTimes, faSkull, faCheck, faUpload} from "@fortawesome/free-solid-svg-icons"

@Component({
  selector: 'app-file-scroller',
  templateUrl: './file-scroller.component.html',
  styleUrls: ['./file-scroller.component.scss']
})
export class FileScrollerComponent implements OnInit {
  @Input() files: Array<OleFile> = [];
  @Output() fileDisplayEvent = new EventEmitter<OleFile>();
  faTimes = faTimes;
  faSkull = faSkull;
  faCheck = faCheck;
  faUpload = faUpload;

  constructor() {
  }

  ngOnInit(): void {
  }

}
