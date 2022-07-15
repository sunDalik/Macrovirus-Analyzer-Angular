import { TestBed } from '@angular/core/testing';

import { VbaAnalysisService } from './ole-analysis.service';

describe('OleAnalysisService', () => {
  let service: VbaAnalysisService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(VbaAnalysisService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
