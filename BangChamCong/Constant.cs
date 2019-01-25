using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BangChamCong
{
    public static class AppSettingConstant
    {
        public static readonly string Author = ConfigurationManager.AppSettings["Author"];

        public static readonly string InOutSheet = ConfigurationManager.AppSettings["InOutSheet"].ToLower();
        public static readonly string EarlyLateSheet = ConfigurationManager.AppSettings["EarlyLateSheet"].ToLower();
        public static readonly string SalarySheet = ConfigurationManager.AppSettings["SalarySheet"].ToLower();
        public static readonly int StartDayColumn = int.Parse(ConfigurationManager.AppSettings["StartDayColumn"]);
        public static readonly int NameColumn = int.Parse(ConfigurationManager.AppSettings["NameColumn"]);

        public static readonly string DateTimeFormat = ConfigurationManager.AppSettings["DateTimeFormat"];
        public static readonly string NormalInTime = ConfigurationManager.AppSettings["NormalInTime"];
        public static readonly string NormalMorningOutTime = ConfigurationManager.AppSettings["NormalMorningOutTime"];
        public static readonly string NormalAfternoonInTime = ConfigurationManager.AppSettings["NormalAfternoonInTime"];
        public static readonly string NormalOutTime = ConfigurationManager.AppSettings["NormalOutTime"];
        public static readonly int LateGapMinute = int.Parse(ConfigurationManager.AppSettings["LateGapMinute"]);
        public static readonly int EarlyGapMinute = int.Parse(ConfigurationManager.AppSettings["EarlyGapMinute"]);
        public static readonly int OTGapMinute = int.Parse(ConfigurationManager.AppSettings["OTGapMinute"]);
        

        public static readonly string[] InOutStart = ConfigurationManager.AppSettings["InOutStart"].ToLower().Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
        public static readonly string InOutEnd = ConfigurationManager.AppSettings["InOutEnd"].ToLower();

        public static readonly string[] ManageList = ConfigurationManager.AppSettings["ManagerList"].ToLower()
                                                                         .Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);

        public static readonly string[] InternList = ConfigurationManager.AppSettings["InternList"].ToLower()
                                                                         .Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
    }
}
