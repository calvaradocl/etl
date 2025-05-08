using System;
using System.Collections.Generic;
using System.Text;

namespace PrevailETL
{
    public class GrafanaQuery
    {
        public string Alias { get; set; }
        public dynamic Query { get; set; }
    }
}
