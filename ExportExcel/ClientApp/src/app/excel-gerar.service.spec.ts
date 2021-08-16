import { TestBed } from '@angular/core/testing';

import { ExcelGerarService } from './excel-gerar.service';

describe('ExcelGerarService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: ExcelGerarService = TestBed.get(ExcelGerarService);
    expect(service).toBeTruthy();
  });
});
