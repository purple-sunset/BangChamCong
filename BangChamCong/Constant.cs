using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BangChamCong
{
    public static class ExcelConstant
    {
        public static readonly string InOutSheet = "BVR";
        public static readonly string EarlyLateSheet = "BDMVS";
        public static readonly string SalarySheet = "BCC";
    }

    public static class AppSettingConstant
    {
        public static readonly string InOutStart = ConfigurationManager.AppSettings["InOutStart"];
        public static readonly string InOutEnd = ConfigurationManager.AppSettings["InOutEnd"];

        public static readonly string[] ManageList = ConfigurationManager.AppSettings["ManagerList"]
                                                                         .Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
    }
}
