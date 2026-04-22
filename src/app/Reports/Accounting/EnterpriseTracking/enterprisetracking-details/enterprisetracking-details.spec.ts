import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EnterprisetrackingDetails } from './enterprisetracking-details';

describe('EnterprisetrackingDetails', () => {
  let component: EnterprisetrackingDetails;
  let fixture: ComponentFixture<EnterprisetrackingDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EnterprisetrackingDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(EnterprisetrackingDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
