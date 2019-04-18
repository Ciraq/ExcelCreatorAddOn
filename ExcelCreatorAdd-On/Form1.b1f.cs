using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Diagnostics;
using ExcelCreatorAdd_On.HelperClasses;

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

            DateTime date1 = Converter.ConvertEditTextValueToDatetime(EditText0.Value.ToString());
            string startDate = date1.ToString("yyyy-MM-dd");

            DateTime date2 = Converter.ConvertEditTextValueToDatetime(EditText1.Value.ToString());
            string endDate = date2.ToString("yyyy-MM-dd");

            string queryHeader = $"call \"Excel_Get_JournalEntry_Header\"('{startDate}', '{endDate}')";
            recordset.DoQuery(queryHeader);

            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string docName = $"{filepath}\\{DateTime.Now.ToString("yyyyMMddHHmmssff")}.xlsx";
            FileInfo sheetinfo = new FileInfo(docName);
            ExcelPackage pck = new ExcelPackage(sheetinfo);

            //Add entryHeader Sheet
            Excel.AddSheet(pck, "entryHeader", recordset);

            string queryDetail = $"call \"Excel_Get_JournalEntry_Detail\"('{startDate}', '{endDate}')";
            recordset.DoQuery(queryDetail);

            //Add entryHeader Sheet
            Excel.AddSheet(pck, "entryDetail", recordset);


            pck.Save();
            Process.Start(docName);
        }

        private SAPbouiCOM.EditText EditText1;
    }
}