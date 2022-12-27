using System;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace GFramework.Xlsx
{
    public class LuaDataModel
    {
        // xlsx文件
        public string xlsx;
        // lua
        public string lua;
        // lua对象类型
        public string type;
        // 命名空间
        public string nameSpace;
        // xlsx sheet
        public DataTable sheet;
        //表明
        public string tblName;
        // lua文本
        public string luaTxt;
    }

    public class SubTbl
    {
        public string sourceNameSpace;
        public string sourceTblName;
        public string sourceDataIndex;

        public string subNameSpace;
        public string subTblName;
        public string subDataIndex;
    }

    public class LuaExporter
    {
        GLogger logger = new GLogger("LuaExporter");

        private List<string> xlsxs = new List<string>();

        // 命名空间、表名
        private Dictionary<string, Dictionary<string, LuaDataModel>> objSheetMap = new Dictionary<string, Dictionary<string, LuaDataModel>>();
        // 命名空间、表名
        private Dictionary<string, Dictionary<string, LuaDataModel>> enumSheetMap = new Dictionary<string, Dictionary<string, LuaDataModel>>();
        // 命名空间、表名
        private Dictionary<string, Dictionary<string, LuaDataModel>> tblSheetMap = new Dictionary<string, Dictionary<string, LuaDataModel>>();

        // 命名空间、表名、索引值、数据项
        private Dictionary<string, Dictionary<string, Dictionary<string, string>>> luaTblMap = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

        // 内嵌表
        private List<SubTbl> subTblList = new List<SubTbl>();

        public void ExportToLua(string sourcePath, string exportPath)
        {
            //----获取所有xlsx文件
            List<string> files = new List<string>();
            this.xlsxs = FileUtil.FindFilesByTypeWithDep(sourcePath, ".xlsx");
            logger.P("扫描到{0}个xlsx文件".Format(this.xlsxs.Count));

            foreach (var xlsx in xlsxs)
            {
                //----忽略不需要导出文件
                if (xlsx.IsPrefixWith(XlsxExporter.Instance.cfg.XlsxIgnoreFlag))
                    continue;
                //----获取所有sheet
                Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (FileStream fs = new FileStream(xlsx, FileMode.Open))
                {
                    string xlsxName = xlsx.TrimSuffix(".");
                    string nameSpace = xlsxName.Suffix(XlsxExporter.Instance.cfg.XlsxNameSpaceFlag);
                    if (string.IsNullOrEmpty(nameSpace)) nameSpace = XlsxExporter.Instance.cfg.LuaDefaultNameSpace;
                    var excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    var set = excelDataReader.AsDataSet().Tables;
                    foreach (DataTable sheet in set)
                    {
                        string sheetName = sheet.TableName;
                        string tblName = sheetName.TrimPrefix(XlsxExporter.Instance.cfg.SheetSepFlag);

                        //----忽略不需要导出的表
                        if (sheetName.IsPrefixWith(XlsxExporter.Instance.cfg.SheetIgnoreFlag))
                            continue;

                        string luaObj = XlsxExporter.Instance.cfg.LuaTypes["Obj"];
                        string luaEnum = XlsxExporter.Instance.cfg.LuaTypes["Enum"];
                        string luaTbl = XlsxExporter.Instance.cfg.LuaTypes["Table"];
                        if (sheetName.IsPrefixWith(luaObj))
                        {
                            LuaDataModel dataTbl = new LuaDataModel()
                            {
                                xlsx = xlsx,
                                type = luaObj,
                                nameSpace = nameSpace,
                                sheet = sheet,
                                tblName = tblName,
                            };
                            if (!this.objSheetMap.ContainsKey(nameSpace)) this.objSheetMap.Add(nameSpace, new Dictionary<string, LuaDataModel>());
                            this.objSheetMap[nameSpace][tblName] = dataTbl;
                        }
                        else if (sheetName.IsPrefixWith(luaEnum))
                        {
                            LuaDataModel dataTbl = new LuaDataModel()
                            {
                                xlsx = xlsx,
                                type = luaObj,
                                nameSpace = nameSpace,
                                sheet = sheet,
                                tblName = tblName,
                            };
                            if (!this.enumSheetMap.ContainsKey(nameSpace)) this.enumSheetMap.Add(nameSpace, new Dictionary<string, LuaDataModel>());
                            this.enumSheetMap[nameSpace][tblName] = dataTbl;
                        }
                        else if (sheetName.IsPrefixWith(luaTbl))
                        {
                            LuaDataModel dataTbl = new LuaDataModel()
                            {
                                xlsx = xlsx,
                                type = luaObj,
                                nameSpace = nameSpace,
                                sheet = sheet,
                                tblName = tblName,
                            };
                            if (!this.tblSheetMap.ContainsKey(nameSpace)) this.tblSheetMap.Add(nameSpace, new Dictionary<string, LuaDataModel>());
                            this.tblSheetMap[nameSpace][tblName] = dataTbl;
                        }
                        else
                        {
                            string prefixFlag = sheetName.Prefix(XlsxExporter.Instance.cfg.SheetSepFlag);
                            throw new NotSupportedException(prefixFlag);
                        }
                    }
                }
            }
            OutputLuaTxt();
            logger.P("导出lua完成！！！");
        }

        private void OutputLuaTxt()
        {
            OutputLuaObj();
            OutputLuaEnum();
            OutputLuaTbl();
            ReplaceSubTblData();
        }

        //-----生成lua 对象
        private void OutputLuaObj()
        {
            foreach (var sheetGroup in this.objSheetMap)
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
                        string fieldValue = ToLuaData(argType, argValue);
                        // 添加注释
                        luaBuilder.AddDesc(argDesc);
                        // 添加字段
                        luaBuilder.AddObjField(argName, fieldValue);

                    }

                    //----lua文件名格式
                    string luaPath = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();
                    dataTbl.lua = luaPath;

                    //----lua导出包
                    string packageName = XlsxExporter.Instance.cfg.Namespace.Format(nameSpace.Upper(), tblName.Upper());
                    string luaData = luaBuilder.ToLocalTbl(packageName);
                    // 添加表说明
                    string tblDesc = LuaBuilder.ToDesc(sheet.Rows[0][0].ToString());
                    string luaBody = tblDesc + luaData;
                    string luaTxt = LuaBuilder.Package(packageName, luaBody);

                    string luaTips = LuaBuilder.ToMultiDesc(dataTbl.xlsx);
                    dataTbl.luaTxt = luaTips + luaTxt;

                    //----写入lua文件
                    WriteToLua(dataTbl);
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
                        string fieldValue = ToLuaData(Xlsx2LuaConfig.Number, argValue);
                        // 添加注释
                        luaBuilder.AddDesc(argDesc);
                        // 添加字段
                        luaBuilder.AddObjField(argName, fieldValue);
                    }

                    //----lua文件名
                    string lua = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();
                    dataTbl.lua = lua;

                    //----lua导出包
                    string packageName = XlsxExporter.Instance.cfg.Namespace.Format(nameSpace.Upper(), tblName.Upper());
                    string luaData = luaBuilder.ToLocalTbl(packageName);
                    // 添加表说明
                    string tblDesc = LuaBuilder.ToDesc(sheet.Rows[0][0].ToString());
                    string luaBody = tblDesc + luaData;
                    string luaTxt = LuaBuilder.Package(packageName, luaBody);


                    //----lua文件描述
                    string luaTips = LuaBuilder.ToMultiDesc(dataTbl.xlsx);
                    dataTbl.luaTxt = luaTips + luaTxt;

                    //----写入lua文件
                    WriteToLua(dataTbl);
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
                    string dataModelName = XlsxExporter.Instance.cfg.LuaDataModel.Format(tblName.Upper());
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

                        for (int j = 1; j < columnsNum; ++j)
                        {
                            // 字段名
                            string filedName = nameList[j];
                            // 字段类型
                            string cellType = typeList[j];

                            // 字段值
                            string cellValue = row[j].ToString();

                            //内嵌表
                            bool isSub = Regex.IsMatch(cellType, XlsxExporter.Instance.cfg.SubTblRegex);
                            if (isSub)
                            {
                                var match = Regex.Match(cellType, XlsxExporter.Instance.cfg.SubTblRegex);
                                string subTblNameSpace = match.Groups[1].Value.ToString().Trim();
                                string subTblName = match.Groups[2].Value.ToString().Trim();
                                SubTbl subTbl = new SubTbl()
                                {
                                    subNameSpace = subTblNameSpace,
                                    subTblName = subTblName,
                                    subDataIndex = cellValue,
                                    sourceNameSpace = nameSpace,
                                    sourceTblName = tblName,
                                    sourceDataIndex = key,
                                };
                                this.subTblList.Add(subTbl);
                            }

                            cellValue = ToLuaData(cellType, cellValue, cellValue);

                            // 添加字段
                            itemBuilder.AddObjField(filedName, cellValue);
                        }

                        string dataItem = LuaBuilder.ToTbl(itemBuilder.ToString());

                        // 添加数据项
                        dataBuilder.AddListItem(key, dataItem);
                        if (!this.luaTblMap.ContainsKey(dataTbl.nameSpace)) this.luaTblMap.Add(dataTbl.nameSpace, new Dictionary<string, Dictionary<string, string>>());
                        if (!this.luaTblMap[dataTbl.nameSpace].ContainsKey(dataTbl.tblName)) this.luaTblMap[dataTbl.nameSpace].Add(dataTbl.tblName, new Dictionary<string, string>());
                        this.luaTblMap[dataTbl.nameSpace][dataTbl.tblName][key] = dataItem;
                    }


                    // lua 文件名
                    string lua = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();
                    dataTbl.lua = lua;

                    //----lua导出包
                    string packageName = XlsxExporter.Instance.cfg.Namespace.Format(nameSpace.Upper(), tblName.Upper());
                    string luaTblData = dataBuilder.ToLocalTbl(packageName);
                    luaBuilder.AddSubBody(luaTblData);

                    string luaBody = luaBuilder.ToString();
                    string luaTxt = LuaBuilder.Package(packageName, luaBody);

                    //----lua文件描述
                    string luaTips = LuaBuilder.ToMultiDesc(dataTbl.xlsx);
                    dataTbl.luaTxt = luaTips + luaTxt;

                    WriteToLua(dataTbl);
                }
            }
        }

        /// <summary>
        /// 替换内嵌表
        /// </summary>
        private void ReplaceSubTblData()
        {
            foreach (var sub in this.subTblList)
            {
                string subNameSpace = sub.subNameSpace;
                string subTblName = sub.subTblName;
                string subDataIndex = sub.subDataIndex;
                string subIndex = "{0}::{1}::{2}".Format(subNameSpace, subTblName, subDataIndex);

                string sourceNameSpace = sub.sourceNameSpace;
                string sourceTblName = sub.sourceTblName;
                string sourceDataIndex = sub.sourceDataIndex;
                string data = this.luaTblMap[subNameSpace][subTblName][subDataIndex];

                var dataModel = this.tblSheetMap[sourceNameSpace][sourceTblName];
                string luaTxt = dataModel.luaTxt;
                dataModel.luaTxt = luaTxt.Replace(subIndex, data);
                WriteToLua(dataModel);
            }
        }

        private string ToLuaData(string cellType, string cellValue, string dataIndex = null)
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
                    var nums = cellValue.Split(XlsxExporter.Instance.cfg.ListSepFlag);
                    StringBuilder numList = new StringBuilder();
                    for (int i = 0; i < nums.Length; ++i)
                    {
                        string v = nums[i].ToString().Trim();
                        if (v.IsDigit())
                        {
                            logger.P("数值数组包含非法字符！！！");
                            continue;
                        }
                        string numItem = LuaTemplate.LIST_ITEM.Format(i, v);
                        numList.AppendLine(numItem);
                    }
                    cellValue = numList.ToString();
                    break;
                case Xlsx2LuaConfig.ListStr:
                    var strs = cellValue.Split(XlsxExporter.Instance.cfg.ListSepFlag);
                    StringBuilder strList = new StringBuilder();
                    for (int i = 0; i < strs.Length; ++i)
                    {
                        string v = strs[i].ToString().Trim();
                        if (v.IsDigit())
                        {
                            logger.P("数值数组包含非法字符！！！");
                            continue;
                        }
                        string strItem = LuaTemplate.LIST_ITEM.Format(i, v);
                        strList.AppendLine(strItem);
                    }
                    cellValue = strList.ToString();
                    break;
                default:
                    bool isSub = Regex.IsMatch(cellType, XlsxExporter.Instance.cfg.SubTblRegex);
                    if (isSub)
                    {
                        var match = Regex.Match(cellType, XlsxExporter.Instance.cfg.SubTblRegex);
                        string subTblNameSpace = match.Groups[1].Value.ToString().Trim();
                        string subTblName = match.Groups[2].Value.ToString().Trim();
                        cellValue = "{0}::{1}::{2}".Format(subTblNameSpace, subTblName, dataIndex);
                    }
                    else
                    {
                        cellValue = string.Empty;
                    }
                    break;
            }
            return cellValue;
        }

        private void WriteToLua(LuaDataModel dataTbl)
        {
            string xlsx = dataTbl.xlsx;
            string lua = dataTbl.lua;
            string luaTxt = dataTbl.luaTxt;
            using (FileStream fs = new FileStream(lua, FileMode.OpenOrCreate, FileAccess.Write))
            {
                fs.SetLength(0);
                byte[] luaData = Encoding.UTF8.GetBytes(luaTxt);
                fs.Write(luaData);
            }
            logger.P("导出完成{0}...".Format(dataTbl.lua));
        }
    }
}