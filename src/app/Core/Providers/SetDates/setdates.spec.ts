import { TestBed } from '@angular/core/testing';

import { Setdates } from './setdates';

describe('Setdates', () => {
  let service: Setdates;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(Setdates);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
