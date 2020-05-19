import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { environment } from 'src/environments/environment';
import { Observable } from 'rxjs/internal/Observable';
import { map } from 'rxjs/operators';

const URL_API = environment.API.EndPoint.NorthWind;

@Injectable({
  providedIn: 'root'
})

export class NorthwindService {

  constructor(private http: HttpClient) { }

  getTop5(dimension: string, orden: string) {
    return this.http.get(`${URL_API}Top5/${dimension}/${orden}`);
  }

  getSerieHistorica(dimension: string, orden: string) {
    return this.http.get(`${URL_API}Histograma/${dimension}/${orden}`);
  }

  getItemsByDimension(dimension: string, orden: string){
    return this.http.get(`${URL_API}GetItemsByDimension/${dimension}/${orden}`).pipe(
      map((result: any) => result.datosDimension)
    )
  }

  getDataPieByDimension(values: string[]){
    console.log('VALUES: ', values);

    return this.http.post(`${URL_API}GetDataPieByDimension`, values).pipe(
      map((result: any) => result)
    )
  }

  getDataBar(values: string[]){
    return this.http.post(`${URL_API}GetDataBar`, values).pipe(
      map((result: any) => result)
    )
  }

}
