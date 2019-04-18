using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace ExcelCreatorAdd_On.HelperClasses
{
    public static class Converter
    {
        public static DateTime ConvertEditTextValueToDatetime(string date)
        {
            try
            {
                SAPbobsCOM.Company company = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
                SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)company.GetBusinessObject(BoObjectTypes.BoBridge);
                DateTime datetime = Convert.ToDateTime(objBridge.Format_StringToDate(date).Fields.Item(0).Value);
                return datetime;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium);
                throw;
            }

        }
    }
}
