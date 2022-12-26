using System.Data;
using System.Text;
using ExcelDataReader;

namespace GFramework.Xlsx
{
    public class LuaDataTbl
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
        // lua文本
        public string luaTxt;
    }

    public class LuaExporter
    {
        GLogger logger = new GLogger("LuaExporter");

        private List<string> xlsxs = new List<string>();

        private Dictionary<string, List<LuaDataTbl>> objSheetMap = new Dictionary<string, List<LuaDataTbl>>();
        private Dictionary<string, List<LuaDataTbl>> enumSheetMap = new Dictionary<string, List<LuaDataTbl>>();
        private Dictionary<string, List<LuaDataTbl>> tblSheetMap = new Dictionary<string, List<LuaDataTbl>>();

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
                    string nameSpace = xlsx.TrimSuffix(".").Suffix(XlsxExporter.Instance.cfg.XlsxNameSpaceFlag);
                    var excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    var set = excelDataReader.AsDataSet().Tables;
                    foreach (DataTable sheet in set)
                    {
                        string tblName = sheet.TableName;
                        //----忽略不需要导出的表
                        if (tblName.IsPrefixWith(XlsxExporter.Instance.cfg.SheetIgnoreFlag))
                            continue;

                        string luaObj = XlsxExporter.Instance.cfg.LuaTypes["Obj"];
                        string luaEnum = XlsxExporter.Instance.cfg.LuaTypes["Enum"];
                        string luaTbl = XlsxExporter.Instance.cfg.LuaTypes["Table"];
                        if (tblName.IsPrefixWith(luaObj))
                        {
                            this.objSheetMap.TryGetValue(nameSpace, out var objs);
                            if (objs == null)
                            {
                                objs = new List<LuaDataTbl>();
                                this.objSheetMap[nameSpace] = objs;
                            }
                            LuaDataTbl dataTbl = new LuaDataTbl()
                            {
                                xlsx = xlsx,
                                type = luaObj,
                                nameSpace = nameSpace,
                                sheet = sheet,
                            };
                            objs.Add(dataTbl);
                        }
                        else if (tblName.IsPrefixWith(luaEnum))
                        {
                            this.enumSheetMap.TryGetValue(nameSpace, out var enums);
                            if (enums == null)
                            {
                                enums = new List<LuaDataTbl>();
                                this.enumSheetMap[nameSpace] = enums;
                            }
                            LuaDataTbl dataTbl = new LuaDataTbl()
                            {
                                xlsx = xlsx,
                                type = luaObj,
                                nameSpace = nameSpace,
                                sheet = sheet,
                            };
                            enums.Add(dataTbl);
                        }
                        else if (tblName.IsPrefixWith(luaTbl))
                        {
                            this.tblSheetMap.TryGetValue(nameSpace, out var tbls);
                            if (tbls == null)
                            {
                                tbls = new List<LuaDataTbl>();
                                this.tblSheetMap[nameSpace] = tbls;
                            }
                            LuaDataTbl dataTbl = new LuaDataTbl()
                            {
                                xlsx = xlsx,
                                type = luaObj,
                                nameSpace = nameSpace,
                                sheet = sheet,
                            };
                            tbls.Add(dataTbl);
                        }
                        else
                        {
                            string prefixFlag = tblName.Prefix(XlsxExporter.Instance.cfg.SheetSepFlag);
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

                //----导出sheet
                foreach (var dataTbl in sheetGroup.Value)
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
                    string tblName = sheet.TableName.TrimPrefix(XlsxExporter.Instance.cfg.SheetSepFlag);
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

                //----导出sheet
                foreach (var dataTbl in sheetGroup.Value)
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
                    string tblName = sheet.TableName.TrimPrefix(XlsxExporter.Instance.cfg.SheetSepFlag);
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

                //----导出sheet
                foreach (var dataTbl in sheetGroup.Value)
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
                    string tblName = sheet.TableName.TrimPrefix(XlsxExporter.Instance.cfg.SheetSepFlag);
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
                            // 字段值
                            string filedName = nameList[j];

                            // 字段值
                            string cellType = typeList[j];
                            string cellValue = row[j].ToString();
                            cellValue = ToLuaData(cellType, cellValue);

                            // 添加字段
                            itemBuilder.AddObjField(filedName, cellValue);
                        }

                        // 添加数据项
                        dataBuilder.AddListItem(key, itemBuilder.ToTbl());
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

        private string ToLuaData(string cellType, string cellValue)
        {
            switch (cellType)
            {
                case Xlsx2LuaConfig.Number:
                    cellValue = LuaTemplate.OBJ.Format(cellValue);
                    break;
                case Xlsx2LuaConfig.String:
                    cellValue = LuaTemplate.STR.Format(cellValue);
                    break;
                case Xlsx2LuaConfig.ListNum:

                    break;
                case Xlsx2LuaConfig.ListStr:

                    break;
            }
            return cellValue;
        }

        private void WriteToLua(LuaDataTbl dataTbl)
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