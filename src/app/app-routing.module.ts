import {NgModule} from '@angular/core';
import {RouterModule, Routes} from '@angular/router';
import {FileUploadComponent} from "./file-upload/file-upload.component";
import {ResultsPageComponent} from "./results-page/results-page.component";

const routes: Routes = [
  {path: 'result', component: ResultsPageComponent, title: "Analysis results"},
  {path: '**', component: FileUploadComponent, title: "Upload a file"},
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule {
}
