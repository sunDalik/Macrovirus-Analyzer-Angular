import { ComponentFixture, TestBed } from '@angular/core/testing';

import { OleFileInputComponent } from './ole-file-input.component';

describe('OleFileInputComponent', () => {
  let component: OleFileInputComponent;
  let fixture: ComponentFixture<OleFileInputComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ OleFileInputComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(OleFileInputComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
