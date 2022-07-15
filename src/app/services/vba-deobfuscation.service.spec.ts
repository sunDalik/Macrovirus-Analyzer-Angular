import { TestBed } from '@angular/core/testing';

import { VbaDeobfuscationService } from './vba-deobfuscation.service';

describe('VbaDeobfuscationService', () => {
  let service: VbaDeobfuscationService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(VbaDeobfuscationService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
