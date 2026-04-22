import { TestBed } from '@angular/core/testing';

import { Sharedservice } from './sharedservice';

describe('Sharedservice', () => {
  let service: Sharedservice;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(Sharedservice);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
