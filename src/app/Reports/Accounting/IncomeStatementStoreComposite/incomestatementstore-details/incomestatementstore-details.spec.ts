import { ComponentFixture, TestBed } from '@angular/core/testing';

import { IncomestatementstoreDetails } from './incomestatementstore-details';

describe('IncomestatementstoreDetails', () => {
  let component: IncomestatementstoreDetails;
  let fixture: ComponentFixture<IncomestatementstoreDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [IncomestatementstoreDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(IncomestatementstoreDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
