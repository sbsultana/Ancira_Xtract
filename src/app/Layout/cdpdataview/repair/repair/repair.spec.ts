import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Repair } from './repair';

describe('Repair', () => {
  let component: Repair;
  let fixture: ComponentFixture<Repair>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [Repair]
    })
    .compileComponents();

    fixture = TestBed.createComponent(Repair);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
