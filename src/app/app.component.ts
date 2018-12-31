import { Component } from '@angular/core';
import * as Highcharts from 'highcharts';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  public Highcharts;
  public chartConstructor;
  public chartOptions: ChartOptions;
  private xlsxFile: File;
  private sheetJsonData: SheetJsonData[];
  public chartData: ChartData[];
  public visibleChart: boolean;
  public showErrorMessage: boolean;
  private xAxisCategories: string[];
  private chartType: string;
  private chartTitle: string;

  constructor() {
    this.Highcharts = Highcharts;
    this.chartConstructor = 'chart';
    this.visibleChart = false;
    this.chartData = [];
    this.xAxisCategories = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    this.chartType = 'scatter';
    this.chartTitle = 'Mini report';
    this.chartOptions = {
      title: {
        text: ''
      },
      series: [{
        data: [],
        type: ''
      }],
      xAxis: {
        categories: []
      }
    };
  }

  public uploadXlsx(file): void {
    this.showErrorMessage = false;
    this.xlsxFile = file[0];
    this.readXlsx();
  }

  private readXlsx(): void {
    const fileReader = new FileReader();
    fileReader.onload = () => {
      const arrayBuffer: any = fileReader.result;
      const data: Uint8Array = new Uint8Array(arrayBuffer);
      const arr: string[] = new Array<string>();
      for (let i = 0; i !== data.length; ++i) {
        arr[i] = String.fromCharCode(data[i]);
      }
      const bstr: string = arr.join('');
      const workbook: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});
      const first_sheet_name: string = workbook.SheetNames[0];
      const worksheet: XLSX.WorkSheet = workbook.Sheets[first_sheet_name];
      this.sheetJsonData = XLSX.utils.sheet_to_json(worksheet, {raw: true});
      this.loadChart();
    };
    fileReader.readAsArrayBuffer(this.xlsxFile);
  }

  private loadChart(): void {
    this.setSeriesData();
    this.setChartType();
    this.setChartTitle();
    this.setChartXAxisCategories();
  }

  private setSeriesData(): void {
    try {
      this.sheetJsonData.map(row => {
        if (!row.Id) {
          this.showErrorMessage = true;
          throw new Error;
        }
        const data: ChartData = {
          id: row.Id,
          name: row.Name,
          x: this.getRandomNumber0To6(),
          y: this.getRandomNumber0To6()
        };
        this.chartData.push(data);
        this.chartOptions.series[0].data.push(data);
      });
      this.visibleChart = true;
    } catch (error) {
      this.showErrorMessage = true;
    }
  }

  private setChartTitle(): void {
    this.chartOptions.title.text = this.chartTitle;
  }

  private setChartType(): void {
    this.chartOptions.series[0].type = this.chartType;
  }

  private setChartXAxisCategories(): void {
    this.chartOptions.xAxis.categories = this.xAxisCategories;
  }

  private getRandomNumber0To6(): number {
    return Math.floor(Math.random() * Math.floor(7));
  }
}

interface SheetJsonData {
  Id: string;
  Name: string;
  __rowNum__: number;
}

interface ChartData {
  id: string;
  name: string;
  x: number;
  y: number;
}

interface ChartOptions {
  title: ChartTitle;
  series: ChartSeriesOptions[];
  xAxis: ChartXAxis;
}

interface ChartTitle {
  text: string;
}

interface ChartSeriesOptions {
  data: ChartData[];
  type: string;
}

interface ChartXAxis {
  categories: string[];
}
