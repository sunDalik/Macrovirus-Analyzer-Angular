import {Component, OnInit} from '@angular/core';
import {Router} from "@angular/router";
import {OleFile} from "../models/ole-file";
import {GlobalDataService} from "../services/global-data.service";

@Component({
  selector: 'app-file-upload',
  templateUrl: './file-upload.component.html',
  styleUrls: ['./file-upload.component.scss']
})
export class FileUploadComponent implements OnInit {

  constructor(private router: Router, private globalData: GlobalDataService) {
  }

  ngOnInit(): void {
  }

  onFilesUpload(files: Array<OleFile>) {
    this.router.navigate(['result']);
    this.globalData.files = files;
  }
}
