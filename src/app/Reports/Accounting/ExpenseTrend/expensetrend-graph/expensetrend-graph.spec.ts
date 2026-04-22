import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExpensetrendGraph } from './expensetrend-graph';

describe('ExpensetrendGraph', () => {
  let component: ExpensetrendGraph;
  let fixture: ComponentFixture<ExpensetrendGraph>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ExpensetrendGraph]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ExpensetrendGraph);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
