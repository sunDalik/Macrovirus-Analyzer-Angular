import { TestBed } from '@angular/core/testing';

import { FileReaderService } from './file-reader.service';

describe('OleFileReaderService', () => {
  let service: FileReaderService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(FileReaderService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
