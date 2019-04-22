using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCreatorAdd_On
{
    public class DIConnection
    {
        private static Company _company = null;
        public static Company company
        {
            get
            {
                if (_company == null)
                {
                    _company = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
                }
                return _company;
            }
        }
    }
}
