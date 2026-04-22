import { ComponentFixture, TestBed } from '@angular/core/testing';

import { IncomestatementtrendGraph } from './incomestatementtrend-graph';

describe('IncomestatementtrendGraph', () => {
  let component: IncomestatementtrendGraph;
  let fixture: ComponentFixture<IncomestatementtrendGraph>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [IncomestatementtrendGraph]
    })
    .compileComponents();

    fixture = TestBed.createComponent(IncomestatementtrendGraph);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
