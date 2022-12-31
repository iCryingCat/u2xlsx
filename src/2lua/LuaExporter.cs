using System.Linq;
using System.Runtime.InteropServices;
using System;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace GFramework.Xlsx
{
    public class LuaDataModel
    {
        // 原xlsx文件路径
        public string xlsx;
        // 导出lua文件路径
        public string luaPath;
        // 命名空间
        public string nameSpace;
        // lua对象类型
        public string luaType;
        //表名
        public string tblName;
        // xlsx sheet
        public DataTable sheet;
        // lua文本
        public string luaTxt;
        // 数据表
        public Dictionary<string, LuaTblData> tblDataMap;

        public string ToLuaIndex()
        {
            return LuaExporter.InlineTblBelongRegexFormat.Format(this.nameSpace, this.tblName);
        }
    }

    public class LuaTblData
    {
        public Dictionary<string, KeyValuePair<string, string>> data = new Dictionary<string, KeyValuePair<string, string>>();

        public LuaTblData()
        {
        }

        public LuaTblData(Dictionary<string, KeyValuePair<string, string>> data)
        {
            this.data = data;
        }

        public override string ToString()
        {
            LuaBuilder itemBuilder = new LuaBuilder();
            foreach (var (name, value) in this.data)
            {
                itemBuilder.AddObjField(name, value.Value);
            }
            return LuaBuilder.ToTbl(itemBuilder.ToString());
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
            return LuaExporter.InlineTblRegexFormat.Format(nameSpace, tblName, dataIndex);
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
            return LuaExporter.InlineTblRegexFormat.Format(nameSpace, tblName, dataIndex);
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

    public class LuaExporter
    {
        GLogger logger = new GLogger("LuaExporter");

        public const string InlineTblRegexFormat = "{0}::{1}::{2}";
        public const string InlineTblBelongRegexFormat = "{0}::{1}";
        public const string InlineTblRegex = "([.*?^[a-zA-Z_][a-zA-Z0-9_]*\\s*::\\s*[a-zA-Z_][a-zA-Z0-9_]*\\s*::\\s*[\\s+a-zA-Z0-9_,]+)";
        public const string InlineTblBelongRegex = "^([a-zA-Z_][a-zA-Z0-9_]*)\\s*::\\s*([a-zA-Z_][a-zA-Z0-9_]*)\\s*::\\s*([\\s+a-zA-Z0-9_,]+)";

        private List<string> xlsxs = new List<string>();

        // 命名空间、表名
        private Dictionary<string, Dictionary<string, LuaDataModel>> objSheetMap = new Dictionary<string, Dictionary<string, LuaDataModel>>();
        // 命名空间、表名
        private Dictionary<string, Dictionary<string, LuaDataModel>> enumSheetMap = new Dictionary<string, Dictionary<string, LuaDataModel>>();
        // 命名空间、表名
        private Dictionary<string, Dictionary<string, LuaDataModel>> tblSheetMap = new Dictionary<string, Dictionary<string, LuaDataModel>>();

        // 内嵌表
        private Dictionary<string, LinkData> linkMap = new Dictionary<string, LinkData>();

        public void ExportToLua(string sourcePath, string exportPath)
        {
            this.FindAllXlsx(sourcePath);
            this.OutputDataModelsForAllTypeOfLua();
            this.OutputLuaTxt();
            this.LinkInlineTbl();
            this.WriteAllXlsxDataToLua();
        }

        private void FindAllXlsx(string sourcePath)
        {
            //----获取所有xlsx文件
            List<string> files = new List<string>();
            this.xlsxs = FileUtil.FindFilesByTypeWithDep(sourcePath, ".xlsx");
            logger.P("扫描到{0}个xlsx文件".Format(this.xlsxs.Count));
        }

        private void OutputDataModelsForAllTypeOfLua()
        {
            foreach (var xlsx in this.xlsxs)
            {
                //----忽略不需要导出文件
                string xlsxFileName = xlsx.GetCurFileName();
                if (Regex.IsMatch(xlsxFileName, XlsxExporter.Instance.cfg.LuaConfig.IgnoreXlsxRegex))
                    continue;

                //----获取所有sheet
                Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (FileStream fs = new FileStream(xlsx, FileMode.Open))
                {
                    var matchNameSpace = Regex.Match(xlsxFileName, XlsxExporter.Instance.cfg.LuaConfig.NameSpaceRegex);
                    string nameSpace = matchNameSpace.Groups[1].Value;

                    //----默认全局命名空间
                    if (string.IsNullOrEmpty(nameSpace))
                        nameSpace = XlsxExporter.Instance.cfg.LuaConfig.LuaDefaultNameSpace;

                    var excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    var sheetSet = excelDataReader.AsDataSet().Tables;
                    foreach (DataTable sheet in sheetSet)
                    {
                        string sheetName = sheet.TableName;

                        //----忽略不需要导出的表
                        if (Regex.IsMatch(sheetName, XlsxExporter.Instance.cfg.LuaConfig.IgnoreSheetRegex))
                            continue;

                        var luaTableMatch = Regex.Match(sheetName, XlsxExporter.Instance.cfg.LuaConfig.SheetRegex);
                        string luaType = luaTableMatch.Groups[1].Value;
                        string tblName = luaTableMatch.Groups[2].Value;

                        LuaDataModel dataTbl = new LuaDataModel();
                        dataTbl.xlsx = xlsx;
                        dataTbl.nameSpace = nameSpace;
                        dataTbl.tblName = tblName;
                        dataTbl.sheet = sheet;

                        dataTbl.luaType = luaType;
                        switch (luaType)
                        {
                            case "Obj":
                                if (!this.objSheetMap.ContainsKey(nameSpace)) this.objSheetMap.Add(nameSpace, new Dictionary<string, LuaDataModel>());
                                this.objSheetMap[nameSpace][tblName] = dataTbl;
                                break;
                            case "Enum":
                                if (!this.enumSheetMap.ContainsKey(nameSpace)) this.enumSheetMap.Add(nameSpace, new Dictionary<string, LuaDataModel>());
                                this.enumSheetMap[nameSpace][tblName] = dataTbl;
                                break;
                            case "Tbl":
                                if (!this.tblSheetMap.ContainsKey(nameSpace)) this.tblSheetMap.Add(nameSpace, new Dictionary<string, LuaDataModel>());
                                this.tblSheetMap[nameSpace][tblName] = dataTbl;
                                break;
                            default:
                                logger.E("不支持导出该lua数据类型：{0}".Format(luaType));
                                break;
                        }
                    }
                }
            }
        }

        private void OutputLuaTxt()
        {
            OutputLuaObj();
            OutputLuaEnum();
            OutputLuaTbl();
        }

        //-----生成lua 对象
        private void OutputLuaObj()
        {
            foreach (var sheetGroup in this.objSheetMap)
            {
                string nameSpace = sheetGroup.Key;

                //----根据命名空间创建目录
                string exportPath = XlsxExporter.Instance.cfg.ExportPath;
                string dirPath = Path.Combine(exportPath, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var sheetMap = sheetGroup.Value.ToArray();
                //----导出sheet
                foreach (var (tblName, dataTbl) in sheetMap)
                {
                    var sheet = dataTbl.sheet;

                    //----检查表格式
                    int columnsNum = sheet.Columns.Count;
                    int rowsNum = sheet.Rows.Count;
                    if (rowsNum < 2 || columnsNum < 4)
                    {
                        logger.E("导出Lua Obj错误： 表内容格式不正确！！！\n"
                        + "example："
                        + "row 1: 该表的描述说明\n"
                        + "row 2: key(参数名) | type(参数类型) | value(参数值) | desc(参数说明)");
                        return;
                    }

                    LuaBuilder luaBuilder = new LuaBuilder();

                    for (int i = 2; i < rowsNum; ++i)
                    {
                        DataRow row = sheet.Rows[i];
                        string argName = row[0].ToString();     // 字段名称
                        string argType = row[1].ToString();     // 字段类型
                        string argValue = row[2].ToString();    // 字段值
                        string argDesc = row[3].ToString();     // 字段表述
                        // TODO:格式检查

                        // 根据数据类型生成lua字段
                        var (cellType, cellValue) = ToLuaData(argType, argValue);
                        // 添加注释
                        luaBuilder.AddDesc(argDesc);
                        // 添加字段
                        luaBuilder.AddObjField(argName, cellValue);

                    }

                    //----lua文件名格式
                    string luaPath = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();
                    dataTbl.luaPath = luaPath;

                    //----lua导出包
                    string packageName = XlsxExporter.Instance.cfg.LuaConfig.DataTableFormat.Format(nameSpace.Upper(), tblName.Upper());
                    string luaData = luaBuilder.ToLocalTbl(packageName);
                    // 添加表说明
                    string tblDesc = LuaBuilder.ToDesc(sheet.Rows[0][0].ToString());
                    string luaBody = tblDesc + luaData;
                    string luaTxt = LuaBuilder.Package(packageName, luaBody);

                    string luaTips = LuaBuilder.ToMultiDesc(dataTbl.xlsx);
                    dataTbl.luaTxt = luaTips + luaTxt;

                    this.objSheetMap[nameSpace][tblName] = dataTbl;
                }
            }
        }

        //-----生成lua 枚举
        private void OutputLuaEnum()
        {
            foreach (var sheetGroup in this.enumSheetMap)
            {
                //----根据命名空间创建目录
                string exportPath = XlsxExporter.Instance.cfg.ExportPath;
                string nameSpace = sheetGroup.Key;
                string dirPath = Path.Combine(exportPath, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var sheetMap = sheetGroup.Value.ToArray();
                //----导出sheet
                foreach (var (tblName, dataTbl) in sheetMap)
                {
                    var sheet = dataTbl.sheet;

                    //----检查表格式
                    int columnsNum = sheet.Columns.Count;
                    int rowsNum = sheet.Rows.Count;
                    if (rowsNum < 2 || columnsNum < 3)
                    {
                        logger.E("导出Lua Enum错误： 表内容格式不正确！！！\n"
                        + "example："
                        + "row 1: 该表的描述说明\n"
                        + "row 2: key(参数名) | value(参数值) | desc(参数说明)");
                        return;
                    }

                    LuaBuilder luaBuilder = new LuaBuilder();

                    for (int i = 2; i < rowsNum; ++i)
                    {
                        DataRow row = sheet.Rows[i];
                        string argName = row[0].ToString();     // 字段名称
                        string argValue = row[1].ToString();    // 字段值
                        string argDesc = row[2].ToString();     // 字段表述
                        // TODO:格式检查

                        // 根据数据类型生成lua字段
                        var (cellType, cellValue) = ToLuaData(Xlsx2LuaConfig.Number, argValue);
                        // 添加注释
                        luaBuilder.AddDesc(argDesc);
                        // 添加字段
                        luaBuilder.AddObjField(argName, cellValue);
                    }

                    //----lua文件名
                    string lua = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();
                    dataTbl.luaPath = lua;

                    //----lua导出包
                    string packageName = XlsxExporter.Instance.cfg.LuaConfig.DataTableFormat.Format(nameSpace.Upper(), tblName.Upper());
                    string luaData = luaBuilder.ToLocalTbl(packageName);
                    // 添加表说明
                    string tblDesc = LuaBuilder.ToDesc(sheet.Rows[0][0].ToString());
                    string luaBody = tblDesc + luaData;
                    string luaTxt = LuaBuilder.Package(packageName, luaBody);


                    //----lua文件描述
                    string luaTips = LuaBuilder.ToMultiDesc(dataTbl.xlsx);
                    dataTbl.luaTxt = luaTips + luaTxt;

                    this.enumSheetMap[nameSpace][tblName] = dataTbl;
                }
            }
        }

        //-----生成lua 数据表
        private void OutputLuaTbl()
        {
            foreach (var sheetGroup in this.tblSheetMap)
            {
                //----根据命名空间创建目录
                string exportPath = XlsxExporter.Instance.cfg.ExportPath;
                string nameSpace = sheetGroup.Key;
                string dirPath = Path.Combine(exportPath, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var sheetMap = sheetGroup.Value.ToArray();
                //----导出sheet
                foreach (var (tblName, dataTbl) in sheetMap)
                {
                    var sheet = dataTbl.sheet;
                    dataTbl.tblDataMap = new Dictionary<string, LuaTblData>();

                    //----检查表格式
                    int columnsNum = sheet.Columns.Count;
                    int rowsNum = sheet.Rows.Count;
                    if (rowsNum < 4 || columnsNum < 2)
                    {
                        logger.E("导出Lua Table错误： 表内容格式不正确！！！\n"
                        + "example："
                        + "row 1: 该表的描述说明\n"
                        + "row 2: key(参数名)\n"
                        + "row 3: type(参数名)\n"
                        + "row 4: desc(参数说明)");
                        return;
                    }

                    //----生成表对象
                    List<string> descList = new List<string>();
                    List<string> typeList = new List<string>();
                    List<string> nameList = new List<string>();

                    DataRow fieldDesc;
                    DataRow fieldTypes;
                    DataRow fieldNames;

                    LuaBuilder objBuilder = new LuaBuilder();
                    LuaBuilder tblBuilder = new LuaBuilder();

                    fieldDesc = sheet.Rows[1];
                    fieldTypes = sheet.Rows[2];
                    fieldNames = sheet.Rows[3];
                    for (int i = 0; i < columnsNum; ++i)
                    {
                        string desc = fieldDesc.IsNull(i) ? string.Empty : fieldDesc[i].ToString();
                        string typeName = fieldTypes.IsNull(i) ? throw new NoNullAllowedException("sheet 缺少字段类型: {0}".Format(sheet.TableName)) : fieldTypes[i].ToString();
                        string fieldName = fieldNames.IsNull(i) ? throw new NoNullAllowedException("sheet 缺少字段名: {0}".Format(sheet.TableName)) : fieldNames[i].ToString();
                        descList.Add(desc);
                        typeList.Add(typeName);
                        nameList.Add(fieldName);
                        objBuilder.AddDesc(desc);
                        objBuilder.AddObjField(fieldName, i.ToString());
                    }

                    LuaBuilder luaBuilder = new LuaBuilder();

                    //----添加表说明
                    string tblDesc = sheet.Rows[0][0].ToString();
                    luaBuilder.AddDesc(tblDesc);

                    //----添加表对象
                    string dataModelName = XlsxExporter.Instance.cfg.LuaConfig.DataTableObjectFormat.Format(tblName.Upper());
                    string luaTblObj = objBuilder.ToLocalTbl(dataModelName);
                    luaBuilder.AddSubBody(luaTblObj);

                    //----生成数据项
                    LuaBuilder dataBuilder = new LuaBuilder();
                    for (int i = 4, enumIndex = 0; i < rowsNum; ++i, ++enumIndex)
                    {
                        DataRow row = sheet.Rows[i];

                        LuaBuilder itemBuilder = new LuaBuilder();

                        // 表索引
                        string key = row[0].ToString();
                        dataTbl.tblDataMap[key] = new LuaTblData();

                        for (int j = 1; j < columnsNum; ++j)
                        {
                            // 字段名
                            string filedName = nameList[j];
                            // 字段类型
                            string cellType = typeList[j];
                            // 字段值
                            string cellValue = row[j].ToString();

                            //内嵌表
                            bool isInline = Regex.IsMatch(cellType, XlsxExporter.Instance.cfg.LuaConfig.LuaTypes.InlineTable);
                            if (isInline)
                            {
                                var match = Regex.Match(cellType, XlsxExporter.Instance.cfg.LuaConfig.LuaTypes.InlineTable);
                                string inlineTblNameSpace = match.Groups[1].Value;
                                string inlineTblName = match.Groups[2].Value;
                                LinkData linkData = new LinkData(
                                    new SourceTbl(nameSpace, tblName, key),
                                    new InlineTbl(inlineTblNameSpace, inlineTblName, cellValue)
                                );
                                this.linkMap[linkData.sourceTbl.ToString()] = linkData;

                            }

                            (cellType, cellValue) = ToLuaData(cellType, cellValue);

                            // 添加字段
                            itemBuilder.AddObjField(filedName, cellValue);
                            dataTbl.tblDataMap[key].data[filedName] = new KeyValuePair<string, string>(cellType, cellValue);
                        }

                        string dataItem = LuaBuilder.ToTbl(itemBuilder.ToString());
                        // 添加数据项
                        dataBuilder.AddListItem(key, dataItem);
                    }

                    // lua 文件名
                    string lua = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();
                    dataTbl.luaPath = lua;

                    //----lua导出包
                    string packageName = XlsxExporter.Instance.cfg.LuaConfig.DataTableFormat.Format(nameSpace.Upper(), tblName.Upper());
                    string luaTblData = dataBuilder.ToLocalTbl(packageName);
                    luaBuilder.AddSubBody(luaTblData);

                    string luaBody = luaBuilder.ToString();
                    string luaTxt = LuaBuilder.Package(packageName, luaBody);

                    //----lua文件描述
                    string luaTips = LuaBuilder.ToMultiDesc(dataTbl.xlsx);
                    dataTbl.luaTxt = luaTips + luaTxt;

                    this.tblSheetMap[nameSpace][tblName] = dataTbl;
                }
            }
        }

        /// <summary>
        /// 替换内嵌表
        /// </summary>
        private void LinkInlineTbl()
        {
            foreach (var (key, linkData) in this.linkMap)
            {
                string inlineTxt = DepLinkInlineTbl(linkData, linkData);
                string inlineTblIndex = linkData.inlineTbl.ToString();

                var sourceTbl = linkData.sourceTbl;
                string sourceNameSpace = sourceTbl.nameSpace;
                string sourceTblName = sourceTbl.tblName;
                var sourceDataModel = this.tblSheetMap[sourceNameSpace][sourceTblName];
                sourceDataModel.luaTxt = sourceDataModel.luaTxt.Replace(inlineTblIndex, inlineTxt);
            }
        }

        private string DepLinkInlineTbl(LinkData rootData, LinkData linkData)
        {
            var sourceTbl = linkData.sourceTbl;
            string sourceNameSpace = sourceTbl.nameSpace;
            string sourceTblName = sourceTbl.tblName;
            string sourceDataIndex = sourceTbl.dataIndex;
            var sourceTblModel = this.tblSheetMap[sourceNameSpace][sourceTblName];
            string luaTxt = sourceTblModel.luaTxt;

            var inlineTbl = linkData.inlineTbl;
            string inlineNameSpace = inlineTbl.nameSpace;
            string inlineTblName = inlineTbl.tblName;
            string inlineDataIndex = inlineTbl.dataIndex;
            var inlineDataIndexList = inlineDataIndex.Split(',');

            string inlineTblIndexDataTxt = inlineDataIndex;
            LuaBuilder inlineTblBuild = new LuaBuilder();
            foreach (var index in inlineDataIndexList)
            {
                var indexData = this.tblSheetMap[inlineNameSpace][inlineTblName].tblDataMap[index];
                var indexDataTxt = indexData.ToString();

                bool hasInline = Regex.IsMatch(indexDataTxt, InlineTblRegex);
                if (hasInline)
                {
                    var matchs = Regex.Matches(indexDataTxt, InlineTblRegex);
                    foreach (Match nextMatch in matchs)
                    {
                        var nextInlineTblIndex = nextMatch.Groups[1].Value;
                        var matchTblIndex3 = Regex.Match(nextInlineTblIndex, InlineTblBelongRegex);
                        var nameSpace = matchTblIndex3.Groups[1].Value;
                        var tblName = matchTblIndex3.Groups[2].Value;
                        var dataIndex = matchTblIndex3.Groups[3].Value;

                        var rootTbl = rootData.sourceTbl;
                        //----循环嵌套
                        if (nameSpace == rootTbl.nameSpace && tblName == rootTbl.tblName)
                        {
                            var tempTblData = new LuaTblData(indexData.data);
                            for (int i = 0; i < tempTblData.data.Count; i++)
                            {
                                var tmp = tempTblData.data.ElementAt(i);
                                if (tmp.Value.Key == InlineTblBelongRegexFormat.Format(nameSpace, tblName))
                                    tempTblData.data.Remove(tmp.Key);
                            }
                            indexDataTxt = tempTblData.ToString();
                            continue;
                        }

                        var nextTbl = rootData.inlineTbl;
                        LinkData inlineLinkData = new LinkData(
                            new SourceTbl(inlineNameSpace, inlineTblName, index),
                            new InlineTbl(nameSpace, tblName, dataIndex)
                        );
                        string inlineTxt = DepLinkInlineTbl(rootData, inlineLinkData);
                        indexDataTxt = indexDataTxt.Replace(dataIndex, inlineTxt);
                    }
                }
                inlineTblBuild.AddListItem(index, indexDataTxt);
            }
            return LuaBuilder.ToTbl(inlineTblBuild.ToString());
        }

        private (string, string) ToLuaData(string cellType, string cellValue)
        {
            cellType = cellType.Trim();
            switch (cellType)
            {
                case Xlsx2LuaConfig.Number:
                    cellValue = LuaTemplate.OBJ.Format(cellValue);
                    break;
                case Xlsx2LuaConfig.String:
                    cellValue = LuaTemplate.STR.Format(cellValue);
                    break;
                case Xlsx2LuaConfig.ListNum:
                    var nums = cellValue.Split(',');
                    LuaBuilder numListBuilder = new LuaBuilder();
                    for (int i = 0; i < nums.Length; ++i)
                    {
                        string v = nums[i].ToString().Trim();
                        if (v.IsDigit())
                        {
                            logger.P("数值数组包含非法字符！！！");
                            continue;
                        }
                        numListBuilder.AddListItem(i.ToString(), v);
                    }
                    cellValue = numListBuilder.ToString();
                    break;
                case Xlsx2LuaConfig.ListStr:
                    var strs = cellValue.Split(',');
                    LuaBuilder strListBuilder = new LuaBuilder();
                    for (int i = 0; i < strs.Length; ++i)
                    {
                        string v = strs[i].ToString().Trim();
                        if (v.IsDigit())
                        {
                            logger.P("数值数组包含非法字符！！！");
                            continue;
                        }
                        strListBuilder.AddListItem(i.ToString(), v);
                    }
                    cellValue = strListBuilder.ToString();
                    break;
                default:
                    bool isinline = Regex.IsMatch(cellType, XlsxExporter.Instance.cfg.LuaConfig.LuaTypes.InlineTable);
                    if (isinline)
                    {
                        var match = Regex.Match(cellType, XlsxExporter.Instance.cfg.LuaConfig.LuaTypes.InlineTable);
                        string inlineTblNameSpace = match.Groups[1].Value;
                        string inlineTblName = match.Groups[2].Value;
                        cellType = InlineTblBelongRegexFormat.Format(inlineTblNameSpace, inlineTblName);
                        cellValue = InlineTblRegexFormat.Format(inlineTblNameSpace, inlineTblName, cellValue);
                    }
                    else
                    {
                        logger.E("配置表数据类型错误！！！");
                    }
                    break;
            }
            return (cellType, cellValue);
        }

        private void WriteAllXlsxDataToLua()
        {
            foreach (var (nameSpace, modelMap) in this.objSheetMap)
            {
                foreach (var (tblName, model) in modelMap)
                {
                    WriteToLua(model);
                }
            }
            foreach (var (nameSpace, modelMap) in this.enumSheetMap)
            {
                foreach (var (tblName, model) in modelMap)
                {
                    WriteToLua(model);
                }
            }

            foreach (var (nameSpace, modelMap) in this.tblSheetMap)
            {
                foreach (var (tblName, model) in modelMap)
                {
                    WriteToLua(model);
                }
            }
        }

        private void WriteToLua(LuaDataModel dataModel)
        {
            string xlsx = dataModel.xlsx;
            string lua = dataModel.luaPath;
            string luaTxt = dataModel.luaTxt;
            using (FileStream fs = new FileStream(lua, FileMode.OpenOrCreate, FileAccess.Write))
            {
                fs.SetLength(0);
                byte[] luaData = Encoding.UTF8.GetBytes(luaTxt);
                fs.Write(luaData);
            }
            logger.P("导出完成{0}...".Format(dataModel.luaPath));
        }
    }
}