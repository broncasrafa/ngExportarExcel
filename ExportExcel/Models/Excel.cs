using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExportExcel.Models
{
    public static class Excel
    {
        public static byte[] Gerar(DataSet ds, string sheetName, string[] cabecalho, string filename)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(sheetName);

            #region Cabeçalho
            int totalColunas = ds.Tables[0].Columns.Count;
            string[] colunasExcel = ExcelMethods.CalculateExcelColumnName(totalColunas);
            string rangeCabecalho = "A1:" + colunasExcel.Last() + "1";
            var range = ws.Range(rangeCabecalho);
            ws.Range(rangeCabecalho).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            ws.Range(rangeCabecalho).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            ws.Range(rangeCabecalho).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            ws.Range(rangeCabecalho).Style.Border.BottomBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
            ws.Range(rangeCabecalho).Style.Border.RightBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
            ws.Range(rangeCabecalho).Style.Border.LeftBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
            ws.Rows(1, 1).AdjustToContents(30.00, 30.00);

            IList<string> nomeColunas = new List<string>();
            cabecalho.ToList().ForEach(coluna => nomeColunas.Add(coluna));

            int index = 0;
            foreach (var col in colunasExcel)
            {
                string rng = col + "1";
                ws.Cell(rng).Value = nomeColunas[index];
                ws.Cell(rng).Style
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                    .Font.SetBold()
                    .Font.SetFontSize(10)
                    .Font.SetFontName("Arial")
                    .Font.SetFontColor(XLColor.Black)
                    .Fill.BackgroundColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#9C9C9C"));

                index++;
            }
            #endregion

            #region Dados
            int inHeaderLenght = 1;
            int inColumn = 0;
            int inRow = 0;

            for (int row = 0; row < ds.Tables[0].Rows.Count; row++)
            {
                for (int column = 0; column < ds.Tables[0].Columns.Count; column++)
                {
                    inColumn = column + 1;
                    inRow = inHeaderLenght + 1 + row; //inicia apartir dessa linha (linha 2 coluna 1 = A2)

                    var valor = ds.Tables[0].Rows[row].ItemArray[column].ToString();

                    ws.Cell(inRow, inColumn).Value = valor;
                    ws.Cell(inRow, inColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    ws.Cell(inRow, inColumn).DataType = ExcelMethods.GetDataTypeOfColumnValue(valor);
                    ws.Cell(inRow, inColumn).Style.NumberFormat.Format = ExcelMethods.FormatCell(valor);

                    ws.Cell(inRow, inColumn).Style
                        .Font.SetFontName("Arial")
                        .Font.SetFontSize(10);
                    ws.Cell(inRow, inColumn).Style.Border.OutsideBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
                    ws.Cell(inRow, inColumn).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    ws.Cell(inRow, inColumn).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    ws.Cell(inRow, inColumn).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    ws.Cell(inRow, inColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                }
            }
            #endregion

            ws.Columns().AdjustToContents();            
            wb.SaveAs(filename);

            byte[] bytes = File.ReadAllBytes(filename);

            //Deletando o arquivo temporario gerado pelo CloseXml
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            return bytes;
        }
        public static byte[] Gerar(DataTable dt, string sheetName, string[] cabecalho, string filename)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add(sheetName);

            #region Cabeçalho
            int totalColunas = dt.Columns.Count;
            string[] colunasExcel = ExcelMethods.CalculateExcelColumnName(totalColunas);
            string rangeCabecalho = "A1:" + colunasExcel.Last() + "1";
            var range = ws.Range(rangeCabecalho);
            ws.Range(rangeCabecalho).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            ws.Range(rangeCabecalho).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            ws.Range(rangeCabecalho).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            ws.Range(rangeCabecalho).Style.Border.BottomBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
            ws.Range(rangeCabecalho).Style.Border.RightBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
            ws.Range(rangeCabecalho).Style.Border.LeftBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
            ws.Rows(1, 1).AdjustToContents(30.00, 30.00);

            IList<string> nomeColunas = new List<string>();
            cabecalho.ToList().ForEach(coluna => nomeColunas.Add(coluna));

            int index = 0;
            foreach (var col in colunasExcel)
            {
                string rng = col + "1";
                ws.Cell(rng).Value = nomeColunas[index];
                ws.Cell(rng).Style
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                    .Font.SetBold()
                    .Font.SetFontSize(10)
                    .Font.SetFontName("Arial")
                    .Font.SetFontColor(XLColor.Black)
                    .Fill.BackgroundColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#9C9C9C"));

                index++;
            }
            #endregion

            #region Dados
            int inHeaderLenght = 1;
            int inColumn = 0;
            int inRow = 0;

            for (int row = 0; row < dt.Rows.Count; row++)
            {
                for (int column = 0; column < dt.Columns.Count; column++)
                {
                    inColumn = column + 1;
                    inRow = inHeaderLenght + 1 + row; //inicia apartir dessa linha (linha 2 coluna 1 = A2)

                    var valor = dt.Rows[row].ItemArray[column].ToString();

                    ws.Cell(inRow, inColumn).Value = valor;
                    ws.Cell(inRow, inColumn).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                    ws.Cell(inRow, inColumn).DataType = ExcelMethods.GetDataTypeOfColumnValue(valor);
                    ws.Cell(inRow, inColumn).Style.NumberFormat.Format = ExcelMethods.FormatCell(valor);

                    ws.Cell(inRow, inColumn).Style
                        .Font.SetFontName("Arial")
                        .Font.SetFontSize(10);
                    ws.Cell(inRow, inColumn).Style.Border.OutsideBorderColor = XLColor.FromColor(System.Drawing.ColorTranslator.FromHtml("#000000"));
                    ws.Cell(inRow, inColumn).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    ws.Cell(inRow, inColumn).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    ws.Cell(inRow, inColumn).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    ws.Cell(inRow, inColumn).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                }
            }
            #endregion

            ws.Columns().AdjustToContents();

            filename = @"C:\Temp\Relatorio_" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".xlsx";
            wb.SaveAs(filename);

            byte[] bytes = File.ReadAllBytes(filename);

            //Deletando o arquivo temporario gerado pelo CloseXml
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            return bytes;
        }
    }
}
