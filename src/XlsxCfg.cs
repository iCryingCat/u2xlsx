namespace GFramework.Xlsx
{
    public class XlsxCfg
    {
        public string SourcePath;
        public string ExportPath;
        public string ExportFlags;
        public string Namespace;
        public string XlsxNameSpaceFlag;
        public string SheetSepFlag;
        public string XlsxIgnoreFlag;
        public string SheetIgnoreFlag;
        public string LuaDataModel;
        public string LuaDefaultNameSpace;
        public string ListSepFlag;
        public string SubTblRegex;

        public Dictionary<string, string> LuaTypes = new Dictionary<string, string>();
    }
}