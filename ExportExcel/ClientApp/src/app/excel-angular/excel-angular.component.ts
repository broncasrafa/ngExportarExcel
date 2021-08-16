import { Component, OnInit } from '@angular/core';
import { Fill, FillPattern, Style, Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { ExcelGerarService } from '../excel-gerar.service';

@Component({
  selector: 'app-excel-angular',
  templateUrl: './excel-angular.component.html',
  styleUrls: ['./excel-angular.component.css']
})
export class ExcelAngularComponent implements OnInit {

  constructor(private excelService: ExcelGerarService) { }

  ngOnInit() {
  }

  gerarExcel() {
    this.excelService.getDados()
      .subscribe((dados: any[]) => {
        console.log(dados);
        this.criarArquivoExcel(dados);
      })
  }


  criarArquivoExcel(dados: any[]) {
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Relatorio');
    
    const headerStyles: Partial<Style> = { 
      font: { name: 'Segoe UI', size: 8, bold: true },
      // fill: {
      //   type: 'pattern', pattern: 'solid', fgColor: { argb: 'BEBEBE' }
      // }
    };
    
    worksheet.columns = [
      { header: 'Id Candidato', key: 'idCandidato', width: 30, style: headerStyles },
      { header: 'Nome', key: 'nome', width: 30, style: headerStyles },
      { header: 'Nome Abreviado', key: 'nomeAbreviado', width: 30, style: headerStyles },
      { header: 'Email Candidato', key: 'emailCandidato', width: 30, style: headerStyles },
      { header: 'Partido', key: 'partido', width: 30, style: headerStyles },
      { header: 'Imagem Candidato', key: 'imagemCandidato', width: 30, style: headerStyles },
      { header: 'Situacao Candidato', key: 'situacaoCandidato', width: 30, style: headerStyles },
      { header: 'Estado', key: 'estado', width: 30, style: headerStyles },
      { header: 'UF', key: 'uf', width: 30, style: headerStyles },
      { header: 'Cidade', key: 'cidade', width: 30, style: headerStyles },
      { header: 'Vice Candidato', key: 'viceCandidato', width: 30, style: headerStyles },
      { header: 'Email Vice Candidato', key: 'emailViceCandidato', width: 30, style: headerStyles },
      { header: 'Imagem Vice Candidato', key: 'imagemViceCandidato', width: 30, style: headerStyles },
      { header: 'Situacao Vice Candidato', key: 'situacaoViceCandidato', width: 30, style: headerStyles },
      { header: 'Data Cadastro', key: 'dataCadastro', width: 30, style: headerStyles }
    ];

    dados.forEach(e => {
      let dataRow = worksheet.addRow({
        idCandidato: e.idCandidato, 
        nome: e.nome, 
        nomeAbreviado: e.nomeAbreviado, 
        emailCandidato: e.emailCandidato, 
        partido: e.partido, 
        imagemCandidato: e.imagemCandidato, 
        situacaoCandidato: e.situacaoCandidato,
        estado: e.estado, 
        uf: e.uf, 
        cidade: e.cidade, 
        viceCandidato: e.viceCandidato, 
        emailViceCandidato: e.emailViceCandidato, 
        imagemViceCandidato: e.imagemViceCandidato, 
        situacaoViceCandidato: e.situacaoViceCandidato, 
        dataCadastro: e.dataCadastro
      });

      let colNumSituacaoCandidato = dataRow.getCell(7);
      if (colNumSituacaoCandidato.value == 'ATIVO') {
        colNumSituacaoCandidato.style = {
          font: { color: { argb: 'FF00FF00' }}
        }
        // colNumSituacaoCandidato..fill = {
        //   type: 'pattern',
        //   pattern: 'solid',
        //   fgColor: { argb: 'ffffff' },
        //   bgColor: { argb: '3D8D25' }
        // }
      } else {
        colNumSituacaoCandidato.style = {
          font: { color: { argb: 'FF0000' }}
        }
      }

      let colNumSituacaoViceCandidato = dataRow.getCell(14);
      if (colNumSituacaoViceCandidato.value == 'ATIVO') {
        colNumSituacaoViceCandidato.style = {
          font: { color: { argb: '3D8D25' }}
        }
      } else {
        colNumSituacaoViceCandidato.style = {
          font: { color: { argb: 'FF0000' }}
        }
      }

      dataRow.eachCell((cell, colNumber) => {
        cell.font = { 
          name: 'Segoe UI', size: 8
        }
      })
    });

    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, `resultado_angular_${new Date().getTime()}.xlsx`);
    })
  }

}
