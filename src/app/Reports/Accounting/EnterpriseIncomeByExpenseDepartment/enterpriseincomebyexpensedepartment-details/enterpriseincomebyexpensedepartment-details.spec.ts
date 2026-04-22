import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EnterpriseincomebyexpensedepartmentDetails } from './enterpriseincomebyexpensedepartment-details';

describe('EnterpriseincomebyexpensedepartmentDetails', () => {
  let component: EnterpriseincomebyexpensedepartmentDetails;
  let fixture: ComponentFixture<EnterpriseincomebyexpensedepartmentDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EnterpriseincomebyexpensedepartmentDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(EnterpriseincomebyexpensedepartmentDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
