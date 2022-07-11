import { TestBed } from '@angular/core/testing';

import { OleFileReaderService } from './ole-file-reader.service';

describe('OleFileReaderService', () => {
  let service: OleFileReaderService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(OleFileReaderService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
