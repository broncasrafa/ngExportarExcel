import { async, ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelCsharpComponent } from './excel-csharp.component';

describe('ExcelCsharpComponent', () => {
  let component: ExcelCsharpComponent;
  let fixture: ComponentFixture<ExcelCsharpComponent>;

  beforeEach(async(() => {
    TestBed.configureTestingModule({
      declarations: [ ExcelCsharpComponent ]
    })
    .compileComponents();
  }));

  beforeEach(() => {
    fixture = TestBed.createComponent(ExcelCsharpComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
