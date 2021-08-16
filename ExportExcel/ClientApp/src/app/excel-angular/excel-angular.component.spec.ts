import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelAngularComponent } from './excel-angular.component';

describe('ExcelAngularComponent', () => {
  let component: ExcelAngularComponent;
  let fixture: ComponentFixture<ExcelAngularComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ExcelAngularComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ExcelAngularComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
