import { ComponentFixture, TestBed } from '@angular/core/testing';

import { FinancialsummaryDetails } from './financialsummary-details';

describe('FinancialsummaryDetails', () => {
  let component: FinancialsummaryDetails;
  let fixture: ComponentFixture<FinancialsummaryDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [FinancialsummaryDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(FinancialsummaryDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
