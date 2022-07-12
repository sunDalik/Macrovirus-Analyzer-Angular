import { ComponentFixture, TestBed } from '@angular/core/testing';

import { FileExplorerWindowComponent } from './file-explorer-window.component';

describe('FileExplorerWindowComponent', () => {
  let component: FileExplorerWindowComponent;
  let fixture: ComponentFixture<FileExplorerWindowComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ FileExplorerWindowComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(FileExplorerWindowComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
