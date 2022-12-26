using System.Data;
using System.Text;
using ExcelDataReader;

namespace GFramework.Xlsx
{
    public class LuaObj
    {

    }

    public class LuaExporter
    {
        GLogger logger = new GLogger("LuaExporter");

        private List<string> xlsxs = new List<string>();

        private Dictionary<string, List<DataTable>> objSheetMap = new Dictionary<string, List<DataTable>>();
        private Dictionary<string, List<DataTable>> enumSheetMap = new Dictionary<string, List<DataTable>>();
        private Dictionary<string, List<DataTable>> tblSheetMap = new Dictionary<string, List<DataTable>>();

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
                                objs = new List<DataTable>();
                                this.objSheetMap[nameSpace] = objs;
                            }
                            objs.Add(sheet);
                        }
                        else if (tblName.IsPrefixWith(luaEnum))
                        {
                            this.objSheetMap.TryGetValue(nameSpace, out var enums);
                            if (enums == null)
                            {
                                enums = new List<DataTable>();
                                this.objSheetMap[nameSpace] = enums;
                            }
                            enums.Add(sheet);
                        }
                        else if (tblName.IsPrefixWith(luaTbl))
                        {
                            this.objSheetMap.TryGetValue(nameSpace, out var tbls);
                            if (tbls == null)
                            {
                                tbls = new List<DataTable>();
                                this.objSheetMap[nameSpace] = tbls;
                            }
                            tbls.Add(sheet);
                        }
                        else
                        {
                            string prefixFlag = tblName.Prefix(XlsxExporter.Instance.cfg.SheetSepFlag);
                            throw new NotSupportedException(prefixFlag);
                        }
                    }
                }
            }
            GenerateLuaObj();
            GenerateLuaEnum();
            GenerateLuaTbl();
            logger.P("导出lua完成！！！");
        }

        //-----生成lua 对象
        private void GenerateLuaObj()
        {
            foreach (var sheetGroup in this.objSheetMap)
            {
                string nameSpace = sheetGroup.Key;
                string exportPath = XlsxExporter.Instance.cfg.ExportPath;
                string dirPath = Path.Combine(exportPath, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);
                foreach (var sheet in sheetGroup.Value)
                {
                    string tblName = sheet.TableName.TrimPrefix(XlsxExporter.Instance.cfg.SheetSepFlag);
                    string luaPath = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();

                    int columnsNum = sheet.Columns.Count;
                    int rowsNum = sheet.Rows.Count;
                    if (rowsNum < 2) throw new Exception("表内容格式不正确！！！\n"
                    + "row1: 该表的描述说明\n"
                    + "row2: key(参数名)、type(参数类型)、value(参数值)、desc(参数说明)");

                    string luaObjName = XlsxExporter.Instance.cfg.Namespace.Format(nameSpace.Upper(), tblName.Upper());

                    LuaBuilder luaBuilder = new LuaBuilder();
                    for (int i = 2; i < rowsNum; ++i)
                    {
                        DataRow row = sheet.Rows[i];
                        string argName = row[0].ToString();
                        string argType = row[1].ToString();
                        string argValue = row[2].ToString();
                        string argDesc = row[3].ToString();
                        luaBuilder.AddDesc(argDesc);
                        string fieldValue = ToLuaData(argType, argValue);
                        luaBuilder.AddObjField(argName, fieldValue);
                    }

                    string luaTxt = luaBuilder.ToLocalTbl(luaObjName);
                    using (FileStream fs = new FileStream(luaPath, FileMode.OpenOrCreate, FileAccess.Write))
                    {
                        byte[] bytes = Encoding.UTF8.GetBytes(luaTxt);
                        fs.Position = 0;
                        fs.Write(bytes, 0, bytes.Length);
                        logger.P("写入字节{0}...".Format(bytes.Length));
                    }
                    logger.P("导出完成{0}...".Format(luaPath));
                }

            }
        }

        private void GenerateLuaEnum()
        {
            foreach (var sheetGroup in this.enumSheetMap)
            {
                string nameSpace = sheetGroup.Key;
                string exportPath = XlsxExporter.Instance.cfg.ExportPath;
                string dirPath = Path.Combine(exportPath, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);
                foreach (var sheet in sheetGroup.Value)
                {
                    string tblName = sheet.TableName.TrimPrefix(XlsxExporter.Instance.cfg.SheetSepFlag);
                    string luaPath = Path.Combine(curDir.FullName, tblName + ".lua").PathFormat();

                    int columnsNum = sheet.Columns.Count;
                    int rowsNum = sheet.Rows.Count;
                    if (rowsNum < 2) throw new Exception("表内容格式不正确！！！\n"
                    + "row1: 该表的描述说明\n"
                    + "row2: key(参数名)、value(参数值)、desc(参数说明)");

                    string luaObjName = XlsxExporter.Instance.cfg.Namespace.Format(nameSpace.Upper(), tblName.Upper());

                    LuaBuilder luaBuilder = new LuaBuilder();
                    for (int i = 2; i < rowsNum; ++i)
                    {
                        DataRow row = sheet.Rows[i];
                        string argName = row[0].ToString();
                        string argValue = row[1].ToString();
                        string argDesc = row[2].ToString();
                        luaBuilder.AddDesc(argDesc);
                        string fieldValue = ToLuaData(argType, argValue);
                        luaBuilder.AddObjField(argName, fieldValue);
                    }

                    string luaTxt = luaBuilder.ToLocalTbl(luaObjName);
                    using (FileStream fs = new FileStream(luaPath, FileMode.OpenOrCreate, FileAccess.Write))
                    {
                        byte[] bytes = Encoding.UTF8.GetBytes(luaTxt);
                        fs.Position = 0;
                        fs.Write(bytes, 0, bytes.Length);
                        logger.P("写入字节{0}...".Format(bytes.Length));
                    }
                    logger.P("导出完成{0}...".Format(luaPath));
                }

            }
        }

        private void GenerateLuaTbl()
        {
            foreach (var sheetGroup in this.tblSheetMap)
            {
                int columnsNum = sheet.Columns.Count;
                int rowsNum = sheet.Rows.Count;
                if (rowsNum < 4) throw new Exception("表内容格式不正确！！！\n"
                + "row1: 该表的描述说明\n" + "row2: 参数说明\n" + "row3: 参数类型\n" + "row3: 参数名\n");

                string sheetName = sheet.TableName;
                LuaBuilder luaBuilder = new LuaBuilder();
                DataRow tblDesc = sheet.Rows[0];
                luaBuilder.AddDesc(tblDesc[0].ToString());

                DataRow fieldDesc;
                DataRow fieldTypes;
                DataRow fieldNames;
                List<string> descList = new List<string>();
                List<string> typeList = new List<string>();
                List<string> nameList = new List<string>();

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
                    objBuilder.AddObjField(fieldName, i.ToString());
                }

                luaBuilder.AddSubContent(objBuilder.ToLocalTbl(tableName));

                for (int i = 4, enumIndex = 0; i < rowsNum; ++i, ++enumIndex)
                {
                    DataRow row = sheet.Rows[i];
                    string key = row[0].ToString();

                    List<string> rowValues = new List<string>();
                    for (int j = 1; j < columnsNum; ++j)
                    {
                        string cellType = typeList[j];
                        string cellValue = row[j].ToString();
                        cellValue = ToLuaData(cellType, cellValue);
                        rowValues.Add(cellValue);
                    }
                    string tblItemKey = row[0].ToString();
                    string tblItemValue = LuaBuilder.ToLuaTable(rowValues.ToArray());
                    tblBuilder.AddListItem(tblItemKey, tblItemValue);
                    break;
                }

                string txt = luaBuilder.ToString();
                using (FileStream fs = new FileStream(luaPath, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(luaTxt);
                    fs.Position = 0;
                    fs.Write(bytes, 0, bytes.Length);
                    logger.P("写入字节{0}...".Format(bytes.Length));
                }
                logger.P("导出完成{0}...".Format(luaPath));
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
}
}