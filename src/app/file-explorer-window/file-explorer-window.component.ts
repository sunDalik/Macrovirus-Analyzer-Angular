import {Component, Input, OnInit} from '@angular/core';
import {OleFile} from "../models/ole-file";

@Component({
  selector: 'app-file-explorer-window',
  templateUrl: './file-explorer-window.component.html',
  styleUrls: ['./file-explorer-window.component.scss']
})
export class FileExplorerWindowComponent implements OnInit {
  @Input() activeFile: OleFile | undefined;
  activeTab: number = 2;

  constructor() {
  }

  ngOnInit(): void {
  }

}
