using System;

namespace BangChamCong
{
    public class Range
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
    }

    public class InOutData
    {
        public DateTime InTime { get; set; }
        public DateTime OutTime { get; set; }
        public double WorkDay { get; set; } = 1;
        public double OTHour { get; set; }
        public string Note { get; set; }
        public string Comment { get; set; }
        public bool IsOTDay { get; set; }
    }
}