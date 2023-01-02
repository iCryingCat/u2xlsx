using Newtonsoft.Json;

namespace GFramework.Xlsx
{
    public class XlsxConfig
    {
        [JsonProperty("Xlsx")]
        public string Xlsx { get; set; }
        [JsonProperty("ExportCmd")]
        public string ExportCmd { get; set; }
        [JsonProperty("LuaConfig")]
        public LuaConfig LuaConfig { get; set; }
        [JsonProperty("JsonConfig")]
        public JsonConfig JsonConfig { get; set; }
        [JsonProperty("XlsxTypeRegex")]
        public string XlsxTypeRegex { get; set; }
        [JsonProperty("DataTableFormat")]
        public string DataTableFormat { get; set; }
        [JsonProperty("DataTableObjectFormat")]
        public string DataTableObjectFormat { get; set; }
        [JsonProperty("LuaDefaultNameSpace")]
        public string LuaDefaultNameSpace { get; set; }
        [JsonProperty("NameSpaceRegex")]
        public string NameSpaceRegex { get; set; }
        [JsonProperty("IgnoreXlsxRegex")]
        public string IgnoreXlsxRegex { get; set; }
        [JsonProperty("IgnoreSheetRegex")]
        public string IgnoreSheetRegex { get; set; }
        [JsonProperty("XlsxTypes")]
        public XlsxTypes XlsxTypes { get; set; }
    }

    public class LuaConfig
    {
        [JsonProperty("ExportTo")]
        public string ExportTo { get; set; }
        [JsonProperty("Externion")]
        public string Externion { get; set; }
        [JsonProperty("DeclareJson")]
        public string DeclareJson { get; set; }
    }

    public class JsonConfig
    {
        [JsonProperty("ExportTo")]
        public string ExportTo { get; set; }
        [JsonProperty("Externion")]
        public string Externion { get; set; }
    }

    public class XlsxTypes
    {
        [JsonProperty("Xid")]
        public string Xid { get; set; }
        [JsonProperty("Number")]
        public string Number { get; set; }
        [JsonProperty("String")]
        public string String { get; set; }
        [JsonProperty("ListNumber")]
        public string ListNumber { get; set; }
        [JsonProperty("ListString")]
        public string ListString { get; set; }
        [JsonProperty("ListSeparator")]
        public string ListSeparator { get; set; }
        [JsonProperty("InlineTable")]
        public string InlineTable { get; set; }

    }

}
