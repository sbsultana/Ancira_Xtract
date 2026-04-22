import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EnterpriseincomebyexpensetrendDetails } from './enterpriseincomebyexpensetrend-details';

describe('EnterpriseincomebyexpensetrendDetails', () => {
  let component: EnterpriseincomebyexpensetrendDetails;
  let fixture: ComponentFixture<EnterpriseincomebyexpensetrendDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EnterpriseincomebyexpensetrendDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(EnterpriseincomebyexpensetrendDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
