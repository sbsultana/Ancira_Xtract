import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SalesgrossDetails } from './salesgross-details';

describe('SalesgrossDetails', () => {
  let component: SalesgrossDetails;
  let fixture: ComponentFixture<SalesgrossDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SalesgrossDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SalesgrossDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
