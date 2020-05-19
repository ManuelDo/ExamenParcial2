import { Component, OnInit } from '@angular/core';
import { NorthwindService } from 'src/app/services/northwind.service';
import { Label } from 'ng2-charts';
import { ChartDataSets } from 'chart.js';
import { Observable } from 'rxjs';

@Component({
  selector: 'app-histograma',
  templateUrl: './histograma.component.html',
  styleUrls: ['./histograma.component.scss']
})
export class HistogramaComponent implements OnInit {

  constructor(private north: NorthwindService) { }

  dataDimension: Label[] = [];
  dataValues: ChartDataSets[] = [];

  DimensionList = [
    { value: 1, label: 'Cliente', dimension: '[Dim Cliente].[Dim Cliente Nombre]' },
    { value: 2, label: 'Producto', dimension: '[Dim Producto].[Dim Producto Nombre]' },
    { value: 3, label: 'Empleado', dimension: '[Dim Empleado].[Dim Empleado Nombre]' },
  ];

  dimensionName$: Observable<any>;
  yearItem$: Observable<any>;
  monthItem$: Observable<any>;

  selectedDimension = null;
  selectedDimensionName: string[] = [];
  selectedYear: string[] = [];
  selectedMonth: string[] = [];

  selectedElements: string[] = [];
  
  ngOnInit(): void {
    this.selectedDimension = this.DimensionList[0];

    this.north.getSerieHistorica(this.selectedDimension.label, 'DESC').subscribe((result: any)=> {
      this.dataDimension = result.datosDimension;

      var dataV = [{
        data: result.datosVenta
      }];

      this.dataValues = dataV;
    });

    this.dimensionName$ = this.north.getItemsByDimension(`[Dim Cliente].[Dim Cliente Nombre]`,'ASC');
    this.yearItem$ = this.north.getItemsByDimension(`[Dim Tiempo].[Dim Tiempo Año]`,'BASC');
    this.monthItem$ = this.north.getItemsByDimension(`[Dim Tiempo].[Dim Tiempo Mes]`,'BASC');
  }

  onChangeDimension($event){
    this.selectedDimension = $event;

    this.north.getSerieHistorica(this.selectedDimension.label, 'DESC').subscribe((result: any)=> {
      this.dataDimension = result.datosDimension;

      var dataV = [{
        data: result.datosVenta
      }];

      this.dataValues = dataV;
    });

    this.dimensionName$ = this.north.getItemsByDimension(`${this.selectedDimension.dimension}`,'ASC');
  }

  onChangeValues($event){
    this.selectedElements = $event;

    this.north.getDataBar(this.selectedElements).subscribe((result: any)=> {
      this.dataDimension = result.datosDimension;

      var dataV = [{
        data: result.datosVenta
      }];

      this.dataValues = dataV;
    });
  }

  clearModel(){
    this.selectedDimensionName = [];
    this.selectedYear = [];
    this.selectedMonth = [];
    this.selectedElements = [];

    
  }
}
