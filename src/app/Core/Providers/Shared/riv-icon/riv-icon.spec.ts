import { ComponentFixture, TestBed } from '@angular/core/testing';

import { RivIcon } from './riv-icon';

describe('RivIcon', () => {
  let component: RivIcon;
  let fixture: ComponentFixture<RivIcon>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [RivIcon]
    })
    .compileComponents();

    fixture = TestBed.createComponent(RivIcon);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
