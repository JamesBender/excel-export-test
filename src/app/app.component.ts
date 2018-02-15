import { Component } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'app';

  doStuff = () => {
    console.log('suff');

    const excelFileName = 'poc_test';
    const data = [
      {name: 'manny', job: 'mechanic'},
      {name: 'manny', job: 'mechanic'},
      {name: 'manny', job: 'mechanic'}
    ];

    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

    const workbook: XLSX.WorkBook = {Sheets: {'data': worksheet}, SheetNames: ['data']};

    const excelBuffer: any = XLSX.write(workbook, {bookType: 'xlsx', type: 'buffer'});

    this.saveAsExcelFile(excelBuffer, excelFileName);
  }

  private saveAsExcelFile(buffer: any, filename: string) {
    const data: Blob = new Blob([buffer], {
      type: EXCEL_TYPE
    });

    FileSaver.saveAs(data, filename + EXCEL_EXTENSION);
  }
}
