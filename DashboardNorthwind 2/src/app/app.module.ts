import { BrowserModule } from '@angular/platform-browser';
import { HttpClientModule } from '@angular/common/http';
import { FormsModule } from '@angular/forms';
import { NgModule } from '@angular/core';

import { ChartsModule } from 'ng2-charts';
import { NgSelectModule } from '@ng-select/ng-select';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';

import { InicioComponent } from './pages/inicio/inicio.component';
import { InformacionComponent } from './pages/info/info.component';
import { PastelComponent } from './pages/pastel/pastel.component';
import { HistogramaComponent } from './pages/histograma/histograma.component';

import { NavbarComponent } from './components/navigation/navbar/navbar.component';
import { BarChartComponent } from './components/charts/bar-chart/bar-chart.component';
import { PieChartComponent } from './components/charts/pie-chart/pie-chart.component';

@NgModule({
  declarations: [
    AppComponent,
    InicioComponent,
    InformacionComponent,
    PastelComponent,
    HistogramaComponent,
    NavbarComponent,
    BarChartComponent,
    PieChartComponent
  ],
  imports: [
    BrowserModule,
    NgSelectModule,
    FormsModule,
    ChartsModule,
    HttpClientModule,
    AppRoutingModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
