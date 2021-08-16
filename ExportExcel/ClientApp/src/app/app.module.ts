import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { RouterModule } from '@angular/router';

import { AppComponent } from './app.component';
import { NavMenuComponent } from './nav-menu/nav-menu.component';
import { HomeComponent } from './home/home.component';
import { ExcelAngularComponent } from './excel-angular/excel-angular.component';
import { ExcelCsharpComponent } from './excel-csharp/excel-csharp.component';
import { ExcelGerarService } from './excel-gerar.service';

@NgModule({
  declarations: [
    AppComponent,
    NavMenuComponent,
    HomeComponent,
    ExcelAngularComponent,
    ExcelCsharpComponent
  ],
  imports: [
    BrowserModule.withServerTransition({ appId: 'ng-cli-universal' }),
    HttpClientModule,
    FormsModule,
    RouterModule.forRoot([
      { path: '', component: HomeComponent, pathMatch: 'full' },
      { path: 'excel-angular', component: ExcelAngularComponent },
      { path: 'excel-csharp', component: ExcelCsharpComponent },
    ])
  ],
  providers: [ ExcelGerarService ],
  bootstrap: [AppComponent]
})
export class AppModule { }
