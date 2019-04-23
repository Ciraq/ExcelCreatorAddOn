using System;
using SAPbouiCOM.Framework;
using System.IO;
using OfficeOpenXml;
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
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
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
            if (String.IsNullOrEmpty(EditText0.Value))
            {
                Application.SBO_Application.SetStatusBarMessage("Zəhmət olmasa tarix əlavə edin");
                return;
            }

            try
            {
                var FirstDayOfMonth = Converter.StringToHanaStyle(Converter.FirstDayOfMonth(EditText0.Value).ToString());
                var LastDayOfMonth = Converter.StringToHanaStyle(Converter.LastDayOfMonth(EditText0.Value).ToString());

                SAPbobsCOM.Recordset recordset1 = (SAPbobsCOM.Recordset)DIConnection.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset recordset2 = (SAPbobsCOM.Recordset)DIConnection.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string queryHeader = $"call \"Excel_Get_JournalEntry_Header\"('{FirstDayOfMonth}', '{LastDayOfMonth}')";
                string queryDetail = $"call \"Excel_Get_JournalEntry_Detail\"('{FirstDayOfMonth}', '{LastDayOfMonth}')";

                var RecordsetHeader = MyRecordset.FillRecordset(queryHeader, recordset1);
                var RecordsetDetail = MyRecordset.FillRecordset(queryDetail, recordset2);

                if (RecordsetHeader.RecordCount==0 && RecordsetDetail.RecordCount==0)
                {
                    Application.SBO_Application.SetStatusBarMessage("Məlumat tapılmadı");
                    return;
                }

                string docName = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\{DateTime.Now.ToString("yyyyMMddHHmmssff")}.xlsx";
                FileInfo sheetinfo = new FileInfo(docName);
                ExcelPackage pck = new ExcelPackage(sheetinfo);

                //Add defterMain
                Excel.AddDefterMainSheet(pck, "defterMain");

                //Add entryHeader Sheet
                Excel.AddSheet(pck, "entryHeader", RecordsetHeader);

                //Add entryHeader Sheet
                Excel.AddSheet(pck, "entryDetail", RecordsetDetail);

                pck.Save();
                Process.Start(docName);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message);
            }
        }

        private SAPbouiCOM.StaticText StaticText0;
    }
}