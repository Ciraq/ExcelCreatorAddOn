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
            //Ayın ilk və son gününü tapıram sonra date formatını dəyişirəm
            string FirstDayOfMonth=string.Empty;
            string LastDayOfMonth=string.Empty;
            if (!String.IsNullOrEmpty(EditText0.Value))
            {
                try
                {
                    FirstDayOfMonth = Converter.StringToHanaStyle(Converter.FirstDayOfMonth(EditText0.Value).ToString());
                    LastDayOfMonth = Converter.StringToHanaStyle(Converter.LastDayOfMonth(EditText0.Value).ToString());
                }
                catch (Exception ex)
                {
                    Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
                    return;
                }
            }
            else
            {
                Application.SBO_Application.SetStatusBarMessage("Zəhmət olmasa tarix əlavə edin", SAPbouiCOM.BoMessageTime.bmt_Medium);
                return;
            }

            SAPbobsCOM.Recordset recordset;
            try
            {
                SAPbobsCOM.Company company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
                throw;
            }  

            try
            {
                string queryHeader = $"call \"Excel_Get_JournalEntry_Header\"('{FirstDayOfMonth}', '{LastDayOfMonth}')";
                recordset.DoQuery(queryHeader);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
            }


            string filepath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string docName = $"{filepath}\\{DateTime.Now.ToString("yyyyMMddHHmmssff")}.xlsx";
            FileInfo sheetinfo = new FileInfo(docName);
            ExcelPackage pck = new ExcelPackage(sheetinfo);

            //Add defterMain
            Excel.AddDefterMainSheet(pck, "defterMain");

            //Add entryHeader Sheet
            Excel.AddSheet(pck, "entryHeader", recordset);

            try
            {
                string queryDetail = $"call \"Excel_Get_JournalEntry_Detail\"('{FirstDayOfMonth}', '{LastDayOfMonth}')";
                recordset.DoQuery(queryDetail);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
            }


            //Add entryHeader Sheet
            Excel.AddSheet(pck, "entryDetail", recordset);


            if (recordset.RecordCount > 0 || recordset.Fields.Count > 0)
            {
                pck.Save();
                Process.Start(docName);
            }

        }

        private SAPbouiCOM.StaticText StaticText0;
    }
}