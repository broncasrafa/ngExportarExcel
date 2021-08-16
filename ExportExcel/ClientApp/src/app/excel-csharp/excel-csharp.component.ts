import { Component, OnInit } from '@angular/core';
import { ExcelGerarService } from '../excel-gerar.service';

@Component({
  selector: 'app-excel-csharp',
  templateUrl: './excel-csharp.component.html',
  styleUrls: ['./excel-csharp.component.css']
})
export class ExcelCsharpComponent implements OnInit {

  constructor(private excelService: ExcelGerarService) { }

  ngOnInit() {
  }

  gerarExcelBlob() {
    this.excelService.gerarExcelCSharp_Blob()
      .subscribe((data: Blob) => {
        this.exportarArquivoUsandoBlob(data, `relatorioBlob_${new Date().getTime()}`);
      })
  }

  gerarExcelBase64() {
    this.excelService.gerarExcelCSharp_Base64()
      .subscribe((result: any) => {
        this.exportarArquivoUsandoBase64String(result.data, `relatorioBase64_${new Date().getTime()}`);
      })
  }

  exportarArquivoUsandoBlob(blobData, nomeArquivo) {
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(blobData);
    link.download = `${nomeArquivo}.xlsx`;
    link.click();
  }

  exportarArquivoUsandoBase64String(base64Str, nomeArquivo) {
    const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';    
    const data = `data:${mimeType};base64,${encodeURIComponent(base64Str)}`;
    const link = document.createElement('a');
    link.href = data;
    link.download = `${nomeArquivo}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

}
