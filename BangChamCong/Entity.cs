using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BangChamCong
{
    public class Range
    {
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int EndRow { get; set; }
        public int EndColumn { get; set; }
    }

    public class InOutData
    {
        public DateTime InTime { get; set; }
        public DateTime OutTime { get; set; }
        public int WorkHour { get; set; }
        public int OTMinute { get; set; }
        public string Note { get; set; }
        public bool IsOTDay { get; set; }
    }
}
