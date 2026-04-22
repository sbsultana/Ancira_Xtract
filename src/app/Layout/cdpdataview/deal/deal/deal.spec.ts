import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Deal } from './deal';

describe('Deal', () => {
  let component: Deal;
  let fixture: ComponentFixture<Deal>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [Deal]
    })
    .compileComponents();

    fixture = TestBed.createComponent(Deal);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
