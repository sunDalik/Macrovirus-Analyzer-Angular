import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { FileUploadComponent } from './file-upload/file-upload.component';
import { ResultsPageComponent } from './results-page/results-page.component';
import { FileScrollerComponent } from './file-scroller/file-scroller.component';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
import { FileExplorerWindowComponent } from './file-explorer-window/file-explorer-window.component';
import { FilePreviewComponent } from './file-preview/file-preview.component';

@NgModule({
  declarations: [
    AppComponent,
    FileUploadComponent,
    ResultsPageComponent,
    FileScrollerComponent,
    FileExplorerWindowComponent,
    FilePreviewComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    FontAwesomeModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
