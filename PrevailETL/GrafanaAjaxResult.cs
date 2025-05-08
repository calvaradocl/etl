using Newtonsoft.Json.Linq;

using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Text;

namespace PrevailETL
{
    public class GrafanaAjaxResult
    {
        public ICollection<JObject> Data { get; set; }
        public string Code { get; set; }
        public string Message { get; set;}
    }

    public class GrafanaInner
    {
        public List<string> aggregateTags { get; set; }
        public Dictionary<string, double> dps { get; set; }
        public string metric { get; set; }
        public Dictionary<string, string> tags { get; set; }
    }

    public class GrafanaResult
    {
        public Dictionary<string, List<GrafanaInner>> data { get; set; }
        public int code { get; set; }
        public string message { get; set; }
    }
}
