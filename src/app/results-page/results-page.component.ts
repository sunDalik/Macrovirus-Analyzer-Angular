import {Component, OnInit} from '@angular/core';
import {Router} from "@angular/router";

@Component({
  selector: 'app-results-page',
  templateUrl: './results-page.component.html',
  styleUrls: ['./results-page.component.scss']
})
export class ResultsPageComponent implements OnInit {

  constructor(private router: Router) {
    const filesData = this.router.getCurrentNavigation()?.extras?.state?.['files'];
    if (filesData == null) router.navigate([''])
  }

  ngOnInit(): void {
  }

}
