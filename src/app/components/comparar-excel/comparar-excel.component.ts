import { Component } from '@angular/core';
import { ExcelService } from 'src/app/services/excel.service';

@Component({
  selector: 'app-comparar-excel',
  templateUrl: './comparar-excel.component.html',
  styleUrls: ['./comparar-excel.component.css']
})
export class CompararExcelComponent {
  archivo1: File | null = null;
  archivo2: File | null = null;
  columnName: string = '';
  resultado: string[][] = [];

  constructor(private excelService: ExcelService) { }

  onFileChange(event: any, fileNumber: number): void {
    const file = event.target.files[0];
    if (fileNumber === 1) {
      this.archivo1 = file;
    } else {
      this.archivo2 = file;
    }
  }

  async compareFiles(): Promise<void> {
    if (this.archivo1 && this.archivo2 && this.columnName) {
      try {
        const data1 = await this.excelService.readExcelFile(this.archivo1);
        const data2 = await this.excelService.readExcelFile(this.archivo2);
        this.resultado = this.excelService.compareDataByColumn(data1, data2, this.columnName);
        this.excelService.exportAsExcel(this.resultado, 'leadsNuevos', 'Hoja1');
      } catch (error: any) {
        alert(error.message);
      }
    } else {
      alert('Por favor, suba ambos archivos y especifique el nombre de la columna.');
    }
  }
}
