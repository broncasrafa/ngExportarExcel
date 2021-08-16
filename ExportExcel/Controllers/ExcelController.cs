using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using ExportExcel.Models;

namespace ExportExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly string[] _cabecalho =
        {
            "Id Candidato",
            "Nome",
            "Nome Abreviado",
            "Email Candidato",
            "Partido",
            "Imagem Candidato",
            "Situacao Candidato",
            "Estado",
            "UF",
            "Cidade",
            "Vice Candidato",
            "Email Vice Candidato",
            "Imagem Vice Candidato",
            "Situacao Vice Candidato",
            "DataCadastro"
        };


        [HttpGet]
        [Route("gerar")]
        public IActionResult GerarExcel()
        {
            try
            {
                IList<Candidato> dados = Candidato.GetCandidatos();
                DataTable dataTable = dados.ToDataTable<Candidato>();
                DataSet dataSet = dados.ToDataSet<Candidato>();

                string filename = $@"C:\Temp\{Guid.NewGuid()}_{DateTime.Now.ToString("ddMMyyyy_HHmmss")}.xlsx";
                byte[] bytesArquivo = Excel.Gerar(dataSet, "relatorio", _cabecalho, filename);
                //byte[] bytesArquivo = Excel.Gerar(dataTable, "relatorio", cabecalho, filename);

                MemoryStream arquivo = new MemoryStream(bytesArquivo);

                return Ok(arquivo);
            }   
            catch (Exception ex)
            {
                return BadRequest($"Um erro ocorreu ao tentar gerar o arquivo: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("dados")]
        public IActionResult ObterDados()
        {
            try
            {
                IList<Candidato> dados = Candidato.GetCandidatos();

                return Ok(dados);
            }
            catch (Exception ex)
            {
                return BadRequest($"Um erro ocorreu ao tentar gerar o arquivo: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("gerar-base64")]
        public IActionResult GerarExcelBase64()
        {
            try
            {
                IList<Candidato> dados = Candidato.GetCandidatos();
                DataTable dataTable = dados.ToDataTable<Candidato>();
                DataSet dataSet = dados.ToDataSet<Candidato>();

                string filename = $@"C:\Temp\{Guid.NewGuid()}_{DateTime.Now.ToString("ddMMyyyy_HHmmss")}.xlsx";
                byte[] bytesArquivo = Excel.Gerar(dataSet, "relatorio", _cabecalho, filename);
                //byte[] bytesArquivo = Excel.Gerar(dataTable, "relatorio", cabecalho, filename);

                MemoryStream arquivo = new MemoryStream(bytesArquivo);

                return Ok(new { data = Convert.ToBase64String(arquivo.ToArray()) });
            }
            catch (Exception ex)
            {
                return BadRequest($"Um erro ocorreu ao tentar gerar o arquivo: {ex.Message}");
            }
        }
    }
}