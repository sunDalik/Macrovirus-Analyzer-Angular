import {Component, Input, OnInit} from '@angular/core';
import {Router} from "@angular/router";
import {OleFile} from "../models/ole-file";

@Component({
  selector: 'app-results-page',
  templateUrl: './results-page.component.html',
  styleUrls: ['./results-page.component.scss']
})
export class ResultsPageComponent implements OnInit {
  files: Array<OleFile> = [];
  activeFile: OleFile;

  constructor(private router: Router) {
    const filesData = this.router.getCurrentNavigation()?.extras?.state?.['files'];
    if (filesData == undefined) router.navigate([''])
    this.files = filesData;
    this.activeFile = this.files[0];
  }

  ngOnInit(): void {
  }

}
