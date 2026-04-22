import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EnterpriseincomebyexpenseDetails } from './enterpriseincomebyexpense-details';

describe('EnterpriseincomebyexpenseDetails', () => {
  let component: EnterpriseincomebyexpenseDetails;
  let fixture: ComponentFixture<EnterpriseincomebyexpenseDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EnterpriseincomebyexpenseDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(EnterpriseincomebyexpenseDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
