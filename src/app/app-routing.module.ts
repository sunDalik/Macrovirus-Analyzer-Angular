import {NgModule} from '@angular/core';
import {RouterModule, Routes} from '@angular/router';
import {FileUploadComponent} from "./file-upload/file-upload.component";
import {ResultsPageComponent} from "./results-page/results-page.component";

const routes: Routes = [
  {path: 'result', component: ResultsPageComponent},
  {path: '**', component: FileUploadComponent},
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule {
}
