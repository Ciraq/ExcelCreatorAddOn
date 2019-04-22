using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace ExcelCreatorAdd_On
{
    public static class Converter
    {
        public static DateTime EditTextToDateTime(string value)
        {
            SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)DIConnection.company.GetBusinessObject(BoObjectTypes.BoBridge);
            DateTime date = Convert.ToDateTime(objBridge.Format_StringToDate(value).Fields.Item(0).Value);
            return date;

            //DateTime CreatdDate = DateTime.ParseExact(date,
            //"yyyyMMdd",
            //System.Globalization.CultureInfo.InvariantCulture);
            //return CreatdDate;
        }

        public static string StringToHanaStyle(string date)
        {
            return Convert.ToDateTime(date).ToString("yyyy-MM-dd");
        }

        public static DateTime FirstDayOfMonth(string date)
        {
            int DayOfMonth = int.Parse(EditTextToDateTime(date).Day.ToString());
            DateTime FirstDayOfMonth = EditTextToDateTime(date).AddDays(-(DayOfMonth - 1));
            return FirstDayOfMonth;
        }

        public static DateTime LastDayOfMonth(string date)
        {
            DateTime LastDayOfMonth = FirstDayOfMonth(date).AddMonths(1).AddDays(-1);
            return LastDayOfMonth;
        }
    }
}
