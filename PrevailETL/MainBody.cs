using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Text;

namespace PrevailETL
{
    public class MainBody
    {
        //[JsonProperty("@odata.context")]
        public string datacontext { get; set; }
        //[JsonProperty("@odata.count")]
        public string datacount { get; set; }
        public ICollection<DataPoint> value { get; set; }
        public string code { get; set; }
        public string message { get; set; }
        public string dataoption { get; set; }
    }
    public class DataPoint
    {
        public int ID { get; set; }
        public string office { get; set; }
        public string company { get; set; }
        public string mine { get; set; }
        public string serialno { get; set; }
        public string equipmenttype { get; set; }
        public string eventcode { get; set; }
        public string eventtype { get; set; }
        public DateTime eventstarttime { get; set; }
        public DateTime inserttime { get; set; }
        public DateTime eventendtime { get; set; }
        public string eventdescription { get; set; }
        public string eventidentifier { get; set; }
        public string sequenceid { get; set; }
        public string message { get; set; }
        public string priority { get; set; }
        public string status { get; set; }
        public string subsystem { get; set; }
        public string component { get; set; }
        public string fromvalue { get; set; }
        public string value { get; set; }
        public string filename { get; set; }
        public string eventgroup { get; set; }
        public string eventgroupno { get; set; }
        public string OperatorID { get; set; }
        public DateTime dataloggertimestamp { get; set; }
    }

    public class DataPointIndex
    {
        public string Index { get; set; }
    }

    public class GrafanaIndex
    {
        public string Index { get; set; }
    }

    public class GrafanaMeasurement
    {
        public string SerialNo { get; set; }
        public DateTime ReadOn { get; set; }
        public string Metric { get; set; }
        public object Value { get; set; }
        public string DataType { get; set; }
        public string DataSource { get; set; }
        public string DataQuality { get; set; }
        public string Alias { get; set; }

    }
}
