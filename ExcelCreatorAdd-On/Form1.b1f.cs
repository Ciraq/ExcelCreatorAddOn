using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Diagnostics;

namespace ExcelCreatorAdd_On
{
    [FormAttribute("ExcelCreatorAdd_On.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Company company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset recordset2 = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string date1 = EditText0.Value;
            string startDate = string.Empty;
            if (date1.Length > 0)
            {
                startDate = date1.Substring(0, 4) + "-" + date1.Substring(4, 2) + "-" + date1.Substring(date1.Length - 2, 2);
            }


            string date2 = EditText1.Value;
            string endDate = string.Empty;
            if (date2.Length > 0)
            {
                endDate = date2.Substring(0, 4) + "-" + date2.Substring(4, 2) + "-" + date2.Substring(date2.Length - 2, 2);
            }

            if (date1.Length > 0 && date2.Length > 0)
            {
                try
                {
                    string queryHeader = $"call \"Excel_Get_JournalEntry_Header\"('{startDate}', '{endDate}')";
                    recordset.DoQuery(queryHeader);
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
                }
            }



            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string docName = $"{filepath}\\{DateTime.Now.ToString("yyyyMMddHHmmssff")}.xlsx";
            FileInfo sheetinfo = new FileInfo(docName);
            ExcelPackage pck = new ExcelPackage(sheetinfo);

            //Add entryHeader Sheet
            Excel.AddSheet(pck, "entryHeader", recordset);

            if (date1.Length > 0 && date2.Length > 0)
            {
                try
                {
                    string queryDetail = $"call \"Excel_Get_JournalEntry_Detail\"('{startDate}', '{endDate}')";
                    recordset.DoQuery(queryDetail);
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
                }
            }


            //Add entryHeader Sheet
            Excel.AddSheet(pck, "entryDetail", recordset);

            if (recordset.RecordCount > 0 || recordset.Fields.Count > 0)
            {
                pck.Save();
                Process.Start(docName);
            }

        }

        private SAPbouiCOM.EditText EditText1;
    }
}