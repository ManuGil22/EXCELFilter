import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  header: any;

  constructor() { }

  readExcelFile(file: File): Promise<any[][]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (event: any) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: any = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        resolve(jsonData);
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  }

  compareDataByColumn(data1: any[][], data2: any[][], columna: string): string[][] {
    const header1 = data1[0];
    const header2 = data2[0];
    this.header = header2;
    const colIndex1 = header1.indexOf(columna);
    const colIndex2 = header2.indexOf(columna);

    if (colIndex1 === -1 || colIndex2 === -1) {
      throw new Error(`La columna "${columna}" no se encuentra en ambos archivos.`);
    }

    const set1 = new Set(data1.slice(1).map(row => row[colIndex1]));
    const set2 = new Set(data2.slice(1).map(row => row[colIndex2]));

    const differenceDeConj = this.setDifference(set2, set1);
    return this.getRows(differenceDeConj, data2, colIndex2);
  }

  private getRows<T>(difference: Set<T>, data2: any[][], colIndex2: number): any[][] {
    const result: any[][] = [];
    for(let e of difference){
        const element = data2.filter(elem => elem[colIndex2] == e);
        result.push(element[0]);
      }
    return result
  }

  private setDifference<T>(setA: Set<T>, setB: Set<T>): Set<T> {
    const difference = new Set<T>();
    setA.forEach(elem => {
      if (!setB.has(elem)) {
        difference.add(elem);
      }
    });
    return difference;
  }

  exportAsExcel(set: string[][], fileName: string, sheetName: string): void {
    const arrayData = [this.header, ...set];
    const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(arrayData);
    const workbook: XLSX.WorkBook = {
      Sheets: { [sheetName]: worksheet },
      SheetNames: [sheetName]
    };

    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  }
}
