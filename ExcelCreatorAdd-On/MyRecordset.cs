using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCreatorAdd_On
{
    public static class MyRecordset
    {
        public static Recordset FillRecordset(string url, Recordset recordset)
        {
            recordset.DoQuery(url);
            return recordset;
        }
    }
}
