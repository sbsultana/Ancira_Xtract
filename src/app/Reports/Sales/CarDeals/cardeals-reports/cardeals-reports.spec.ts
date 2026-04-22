import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CardealsReports } from './cardeals-reports';

describe('CardealsReports', () => {
  let component: CardealsReports;
  let fixture: ComponentFixture<CardealsReports>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [CardealsReports]
    })
    .compileComponents();

    fixture = TestBed.createComponent(CardealsReports);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
