import { ComponentFixture, TestBed } from '@angular/core/testing';

import { EnterpriseincomebyexpensewithoutoverheadDetails } from './enterpriseincomebyexpensewithoutoverhead-details';

describe('EnterpriseincomebyexpensewithoutoverheadDetails', () => {
  let component: EnterpriseincomebyexpensewithoutoverheadDetails;
  let fixture: ComponentFixture<EnterpriseincomebyexpensewithoutoverheadDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [EnterpriseincomebyexpensewithoutoverheadDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(EnterpriseincomebyexpensewithoutoverheadDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
