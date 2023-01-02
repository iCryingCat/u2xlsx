using System.Data;

namespace GFramework.Xlsx
{
    public class XlsxDataModel
    {
        // 原xlsx文件路径
        public string xlsx;
        // 导出文件路径
        public string export;
        // 命名空间
        public string nameSpace;
        // lua对象类型
        public string xlsxType;
        // 表注释
        public string desc;
        // 表名
        public string tblName;
        // lua文本
        public string txt;
        // 数据表
        public List<XlsxTblItemData> objDataList = new List<XlsxTblItemData>();
        // 数据表
        public Dictionary<string, XlsxTblData> tblDataMap = new Dictionary<string, XlsxTblData>();

        public string ToLuaIndex()
        {
            return XlsxExporter.InlineTblBelongRegexFormat.Format(this.nameSpace, this.tblName);
        }
    }

    public class XlsxTblData
    {
        public Dictionary<string, XlsxTblItemData> data = new Dictionary<string, XlsxTblItemData>();

        public XlsxTblData() { }

        public XlsxTblData(Dictionary<string, XlsxTblItemData> data)
        {
            this.data = data;
        }
    }

    public class XlsxTblItemData
    {
        public string fieldType;
        public string fieldName;
        public string fieldValue;
        public string fieldDesc;
        public string group;

        public XlsxTblItemData(string fieldType, string fieldName, string fieldValue, string fieldDesc = null)
        {
            this.fieldType = fieldType;
            this.fieldName = fieldName;
            this.fieldValue = fieldValue;
            this.fieldDesc = fieldDesc;
        }
    }

    public class InlineTbl
    {
        public string nameSpace;
        public string tblName;
        public string dataIndex;

        public InlineTbl(string nameSpace, string tblName, string dataIndex)
        {
            this.nameSpace = nameSpace;
            this.tblName = tblName;
            this.dataIndex = dataIndex;
        }

        public override string ToString()
        {
            return XlsxExporter.InlineTblRegexFormat.Format(nameSpace, tblName, dataIndex);
        }
    }

    public class SourceTbl
    {
        public string nameSpace;
        public string tblName;
        public string dataIndex;

        public SourceTbl(string nameSpace, string tblName, string dataIndex)
        {
            this.nameSpace = nameSpace;
            this.tblName = tblName;
            this.dataIndex = dataIndex;
        }

        public SourceTbl(InlineTbl inlineTbl)
        {
            this.nameSpace = inlineTbl.nameSpace;
            this.tblName = inlineTbl.tblName;
            this.dataIndex = inlineTbl.dataIndex;
        }

        public override string ToString()
        {
            return XlsxExporter.InlineTblRegexFormat.Format(nameSpace, tblName, dataIndex);
        }
    }

    public class LinkData
    {
        public SourceTbl sourceTbl;
        public InlineTbl inlineTbl;

        public LinkData(SourceTbl sourceTbl, InlineTbl inlineTbl)
        {
            this.sourceTbl = sourceTbl;
            this.inlineTbl = inlineTbl;
        }
    }
}