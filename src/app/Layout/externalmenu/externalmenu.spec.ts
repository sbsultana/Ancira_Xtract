import { ComponentFixture, TestBed } from '@angular/core/testing';

import { Externalmenu } from './externalmenu';

describe('Externalmenu', () => {
  let component: Externalmenu;
  let fixture: ComponentFixture<Externalmenu>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [Externalmenu]
    })
    .compileComponents();

    fixture = TestBed.createComponent(Externalmenu);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
