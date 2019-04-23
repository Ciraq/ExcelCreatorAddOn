using OfficeOpenXml;
using OfficeOpenXml.Style;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCreatorAdd_On
{
    public static class Excel
    {
        public static void AddDefterMainSheet(ExcelPackage pck, string sheetName, string htmlColor = "#D3D3D3")
        {
            var list = new List<DefterMain>()
            {
                new DefterMain() {vkn=9999999, periodStart="1/1/2014", periodEnd="1/31/2014", subeKodu=2}
            };
            //Add Column
            var sheet = pck.Workbook.Worksheets.Add(sheetName);
            sheet.Cells["A1"].Value = "vkn";
            sheet.Cells["B1"].Value = "Period_start";
            sheet.Cells["C1"].Value = "Period_end";
            sheet.Cells["D1"].Value = "Sube_Kodu";

            //Fill Cells
            int currentRow = 2;
            foreach (var item in list)
            {
                sheet.Cells["A" + currentRow.ToString()].Value = item.vkn;
                sheet.Cells["B" + currentRow.ToString()].Value = item.periodStart;
                sheet.Cells["C" + currentRow.ToString()].Value = item.periodEnd;
                sheet.Cells["D" + currentRow.ToString()].Value = item.subeKodu;
            }


            //Add Style
            Color headercolor = ColorTranslator.FromHtml(htmlColor);
            sheet.Cells["A1:D1"].Style.Font.Bold = true;
            sheet.Cells["A1:D1"].Style.Font.Bold = true;
            sheet.Cells["A1:D1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells["A1:D1"].Style.Fill.BackgroundColor.SetColor(headercolor);
            sheet.Cells.AutoFitColumns();
            sheet.DefaultColWidth = 18;
            sheet.DefaultRowHeight = 18;
        }

        public static void AddSheet(ExcelPackage pck, string sheetName, Recordset recordset, string htmlColor = "#D3D3D3")
        {
            //recordset.DoQuery(url);

            var sheet = pck.Workbook.Worksheets.Add(sheetName);
            
            //Add Columns
            int col = 1;
            for (int i = 0; i < recordset.Fields.Count; i++)
            {
                var recordCol = recordset.Fields.Item(i).Description;
                sheet.Cells[1, col].Value = recordCol;
                col++;
            }

            //Add Style
            Color headercolor = ColorTranslator.FromHtml(htmlColor);
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Font.Bold = true;
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Font.Bold = true;
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Fill.BackgroundColor.SetColor(headercolor);
            sheet.Cells.AutoFitColumns();
            sheet.DefaultColWidth = 18;
            sheet.DefaultRowHeight = 18;


            //Fill cells
            if (recordset.Fields.Count > 0)
            {
                int currentCol = 1;
                int currentRow = 2;
                for (int i = 0; i < recordset.RecordCount; i++)
                {
                    for (int j = 0; j < recordset.Fields.Count; j++)
                    {
                        var recordCol = recordset.Fields.Item(j).Description;
                        sheet.Cells[currentRow, currentCol].Value = recordset.Fields.Item(recordCol).Value.ToString();
                        currentCol++;
                        if (currentCol > recordset.Fields.Count)
                            currentCol = 1;
                    }
                    currentRow++;
                    recordset.MoveNext();
                }
            }
        }
    }
}
