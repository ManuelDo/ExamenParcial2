import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';

import { InicioComponent } from './pages/inicio/inicio.component';
import { InformacionComponent } from './pages/info/info.component';
import { PastelComponent } from './pages/pastel/pastel.component';
import { HistogramaComponent } from './pages/histograma/histograma.component';

const routes: Routes = [
  { path: 'inicio', component: InicioComponent },
  { path: 'informacion', component: InformacionComponent },
  { path: 'pastel', component: PastelComponent },
  { path: 'histograma', component: HistogramaComponent }
];

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
