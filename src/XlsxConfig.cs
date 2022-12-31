using Newtonsoft.Json;

namespace GFramework.Xlsx
{
    public class XlsxConfig
    {
        [JsonProperty("SourcePath")]
        public string SourcePath { get; set; }
        [JsonProperty("ExportPath")]
        public string ExportPath { get; set; }
        [JsonProperty("ExportFlags")]
        public string ExportFlags { get; set; }
        [JsonProperty("LuaConfig")]
        public LuaConfig LuaConfig { get; set; }

    }

    public class LuaConfig
    {
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
        [JsonProperty("SheetRegex")]
        public string SheetRegex { get; set; }
        [JsonProperty("LuaTypes")]
        public LuaTypes LuaTypes { get; set; }
    }

    public class LuaTypes
    {
        [JsonProperty("Number")]
        public string Number { get; set; }
        [JsonProperty("String")]
        public string String { get; set; }
        [JsonProperty("ListNumber")]
        public string ListNumber { get; set; }
        [JsonProperty("ListString")]
        public string ListString { get; set; }
        [JsonProperty("InlineTable")]
        public string InlineTable { get; set; }
    }
}