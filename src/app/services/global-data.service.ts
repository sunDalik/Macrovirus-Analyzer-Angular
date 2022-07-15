import {Injectable} from '@angular/core';
import {OleFile} from "../models/ole-file";

@Injectable({
  providedIn: 'root'
})
export class GlobalDataService {
  files: Array<OleFile> = [];

  constructor() {
  }
}
