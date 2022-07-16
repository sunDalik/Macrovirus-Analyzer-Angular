import { TestBed } from '@angular/core/testing';

import { PcodeDisassemblyService } from './pcode-disassembly.service';

describe('PcodeDisassemblyService', () => {
  let service: PcodeDisassemblyService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(PcodeDisassemblyService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
