import { ComponentFixture, TestBed } from '@angular/core/testing';

import { BudgetForecastInputVariables } from './budget-forecast-input-variables';

describe('BudgetForecastInputVariables', () => {
  let component: BudgetForecastInputVariables;
  let fixture: ComponentFixture<BudgetForecastInputVariables>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [BudgetForecastInputVariables]
    })
    .compileComponents();

    fixture = TestBed.createComponent(BudgetForecastInputVariables);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
