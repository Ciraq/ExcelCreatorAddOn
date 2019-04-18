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
        public static void AddSheet(ExcelPackage pck, string sheetName, Recordset recordset ,string htmlColor = "#D3D3D3")
        {
            var sheet = pck.Workbook.Worksheets.Add(sheetName);

            Color headercolor = ColorTranslator.FromHtml(htmlColor);
            //Add Columns
            int col = 1;
            for (int i = 0; i < recordset.Fields.Count; i++)
            {
                var recordCol = recordset.Fields.Item(i).Description;
                sheet.Cells[1, col].Value = recordCol;
                col++;
            }

            //Add Style
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Font.Bold = true;
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Font.Bold = true;
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, 1, 1, recordset.Fields.Count].Style.Fill.BackgroundColor.SetColor(headercolor);
            sheet.Cells.AutoFitColumns();
            sheet.DefaultColWidth = 18;
            sheet.DefaultRowHeight = 18;

            //Fill cells
            if (recordset.RecordCount > 0)
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
