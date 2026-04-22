import { ComponentFixture, TestBed } from '@angular/core/testing';

import { IncomeBudgetVariables } from './income-budget-variables';

describe('IncomeBudgetVariables', () => {
  let component: IncomeBudgetVariables;
  let fixture: ComponentFixture<IncomeBudgetVariables>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [IncomeBudgetVariables]
    })
    .compileComponents();

    fixture = TestBed.createComponent(IncomeBudgetVariables);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
