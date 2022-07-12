import { ComponentFixture, TestBed } from '@angular/core/testing';

import { FileScrollerComponent } from './file-scroller.component';

describe('FileScrollerComponent', () => {
  let component: FileScrollerComponent;
  let fixture: ComponentFixture<FileScrollerComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ FileScrollerComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(FileScrollerComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
