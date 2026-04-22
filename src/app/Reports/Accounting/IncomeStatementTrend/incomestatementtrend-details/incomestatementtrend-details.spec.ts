import { ComponentFixture, TestBed } from '@angular/core/testing';

import { IncomestatementtrendDetails } from './incomestatementtrend-details';

describe('IncomestatementtrendDetails', () => {
  let component: IncomestatementtrendDetails;
  let fixture: ComponentFixture<IncomestatementtrendDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [IncomestatementtrendDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(IncomestatementtrendDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
