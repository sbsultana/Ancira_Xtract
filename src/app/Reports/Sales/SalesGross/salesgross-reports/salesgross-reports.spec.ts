import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SalesgrossReports } from './salesgross-reports';

describe('SalesgrossReports', () => {
  let component: SalesgrossReports;
  let fixture: ComponentFixture<SalesgrossReports>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SalesgrossReports]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SalesgrossReports);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
