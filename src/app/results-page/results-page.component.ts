import {Component, Input, OnInit} from '@angular/core';
import {Router} from "@angular/router";
import {OleFile} from "../models/ole-file";
import {GlobalDataService} from "../services/global-data.service";

@Component({
  selector: 'app-results-page',
  templateUrl: './results-page.component.html',
  styleUrls: ['./results-page.component.scss']
})
export class ResultsPageComponent implements OnInit {
  files: Array<OleFile> = [];
  activeFile: OleFile;

  constructor(private router: Router, private globalData: GlobalDataService) {
    const filesData = this.globalData.files;
    if (filesData == undefined || filesData.length === 0) router.navigate([''])
    this.files = filesData;
    this.activeFile = this.files[0];
  }

  ngOnInit(): void {
  }

}
