using System;
using System.Collections.Generic;
using System.Text;

namespace PrevailETL
{
    public class ControlPrevail

    {
        public string ProcessType { get; set; }
        public bool WasSuccessfull { get; set; }
        public string ProcessCode { get; set; }
        public int DataCount { get; set; }
        public float ProcessTime { get; set; }
        public DateTime Day { get; set; }
        public DateTime DayHour { get; set; }
        public DateTime TimeStamp { get; set; }
        public string Message { get; set; }
    }
}
