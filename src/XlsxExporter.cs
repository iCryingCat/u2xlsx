using System.Data;
using System.Text;

namespace GFramework.Xlsx
{
    /// <summary>
    /// xlsx导出工具
    /// </summary>
    public class XlsxExporter : Singleton<XlsxExporter>
    {
        public XlsxConfig cfg = null;

        public XlsxConfig LoadConfig()
        {
            string curDir = Environment.CurrentDirectory;
            string cfgName = "xlsx.config.json";
            string cfgPath = Path.Combine(curDir, cfgName).Format();
            JsonStream js = new JsonStream(cfgPath);
            var cfg = js.Read<XlsxConfig>();
            this.cfg = cfg;
            this.cfg.SourcePath = Path.GetFullPath(cfg.SourcePath).PathFormat();
            this.cfg.ExportPath = Path.GetFullPath(cfg.ExportPath).PathFormat();
            return cfg;
        }

        // private class TableProperty
        // {
        //     public string nameSpace;
        //     public string fieldName;
        //     public string fieldValue;

        //     public TableProperty(string nameSpace, string fieldName, string fieldValue)
        //     {
        //         this.nameSpace = nameSpace;
        //         this.fieldName = fieldName;
        //         this.fieldValue = fieldValue;
        //     }
        // }

        // private static Dictionary<string, TableProperty> tableMap = new Dictionary<string, TableProperty>();
        // private const int dfsMaxDepth = 50;
        // private static int dfsDepth = 0;

        // //----加载配置文件
        // static private Dictionary<string, string> LoadConfig()
        // {
        //     string curDir = Environment.CurrentDirectory;
        //     if (curDir == null) throw new FileNotFoundException("XlsxExporterEditor");

        //     string cfgName = "xlsx.config.json";
        //     string cfgPath = Path.Combine(curDir, cfgName).Format();
        //     JsonStream js = new JsonStream(cfgPath);
        //     Dictionary<string, string> cfg = js.Read<Dictionary<string, string>>();
        //     return cfg;
        // }



        // //----导出cs
        // static public void ExecuteExportToCS(string rootPath)
        // {
        //     List<string> files = new List<string>();
        //     GetAllXlsxFile(rootPath, ref files);
        //     foreach (var file in files)
        //     {
        //         ReadXlsx(file, XlsxConfig.DOT_CS);
        //     }
        // }

        // static private void GetAllXlsxFile(string rootPath, ref List<string> filesArr)
        // {
        //     var directories = Directory.GetDirectories(rootPath);
        //     var files = Directory.GetFiles(rootPath);
        //     foreach (var directory in directories)
        //     {
        //         GetAllXlsxFile(directory, ref filesArr);
        //     }
        //     foreach (var file in files)
        //     {
        //         string extension = file.PathExtension();
        //         if (extension == XlsxConfig.DOT_XLSX)
        //             filesArr.Add(file);
        //     }
        // }

        // static private void ReadXlsx(string file, string toBuildFileExtension)
        // {
        //     using (FileStream xlsxFS = new FileStream(file, FileMode.Open))
        //     {
        //         IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(xlsxFS);
        //         DataSet set = excelDataReader.AsDataSet();
        //         foreach (DataTable sheet in set.Tables)
        //         {
        //             string sheetName = sheet.TableName;
        //             // 是否忽略表
        //             if (sheetName[0] == XlsxConfig.KEY_IGNORE) continue;

        //             string[] flags = sheetName.Split(XlsxConfig.KEY_SHEET_Sep);
        //             if (flags.Length < 2)
        //                 throw new Exception("xlsx sheet：{0}命名不合法！！！ 应该至少标明表类型、表名，比如：Obj-test");

        //             // 表类型
        //             string sheetFlag = flags[0];
        //             XlsxSheetFlag xlsxSheetFlag;
        //             if (Enum.TryParse<XlsxSheetFlag>(sheetFlag, out xlsxSheetFlag))
        //                 throw new Exception("无法导出该格式的xlsx sheet：{0}".Format(sheetFlag));

        //             // 表名
        //             string tableName = flags[1];

        //             // 命名空间
        //             string nameSpace = flags[2];

        //             // 导出路径
        //             string buildFileName = tableName + toBuildFileExtension;
        //             string toBuildPath = Path.Combine(XlsxConfig.BUILD_PATH, buildFileName);

        //             TableProperty property;
        //             if (tableMap.TryGetValue(buildFileName, out property)) continue;
        //             string txt = string.Empty;
        //             switch (xlsxSheetFlag)
        //             {
        //                 case XlsxSheetFlag.Class:
        //                     txt = BuildCSClass(nameSpace, tableName, sheet);
        //                     break;
        //                 case XlsxSheetFlag.Obj:
        //                     txt = BuildLuaObj(nameSpace, tableName, sheet);
        //                     break;
        //                 case XlsxSheetFlag.Enum:
        //                     txt = BuildLuaEnum(nameSpace, tableName, sheet);
        //                     break;
        //                 case XlsxSheetFlag.Tbl:
        //                     txt = BuildLuaTable(nameSpace, tableName, sheet);
        //                     break;
        //             }

        //             using (FileStream fs = new FileStream(toBuildPath, FileMode.OpenOrCreate))
        //             {
        //                 byte[] bytes = Encoding.UTF8.GetBytes(txt);
        //                 fs.Write(bytes, 0, bytes.Length);
        //             }
        //         }
        //     }
        // }

        // static private string BuildCSClass(string nameSpace, string className, DataTable sheet)
        // {
        //     int colsNum = sheet.Columns.Count;
        //     int rowsNum = sheet.Rows.Count;
        //     if (rowsNum < 4) throw new Exception("表内容格式不正确！！！1.该表的描述说明、2.字段描述、3.字段类型、4.字段名");

        //     CSBuilder csBuilder = new CSBuilder(nameSpace);
        //     DataRow typeDesc = sheet.Rows[0];
        //     DataRow fieldDesc = sheet.Rows[1];
        //     DataRow fieldTypes = sheet.Rows[2];
        //     DataRow fieldNames = sheet.Rows[3];
        //     // DataModel 类
        //     Dictionary<Type, string> namespaceMap = new Dictionary<Type, string>();
        //     csBuilder.desc = typeDesc.IsNull(0) ? typeDesc[0].ToString() : string.Empty;
        //     for (int i = 0; i < colsNum; ++i)
        //     {
        //         string desc = fieldDesc.IsNull(i) ? string.Empty : fieldDesc[i].ToString();
        //         string typeName = fieldTypes.IsNull(i) ? throw new NoNullAllowedException("sheet 缺少字段类型: {0}".Format(sheet.TableName)) : fieldTypes[i].ToString();
        //         string fieldName = fieldNames.IsNull(i) ? throw new NoNullAllowedException("sheet 缺少字段名: {0}".Format(sheet.TableName)) : fieldNames[i].ToString();
        //         csBuilder.AddDesc(desc);
        //         csBuilder.AddPublicField(typeName, fieldName);
        //         Type t = Type.GetType(fieldName);
        //         if (null != t)
        //         {
        //             string ns = t.Namespace;
        //             if (namespaceMap.TryGetValue(t, out ns))
        //             {
        //                 csBuilder.AddUsing(ns);
        //             }
        //         }
        //     }
        //     csBuilder.AddSubClass(className, typeof(XlsxModel).Name);
        //     string txt = csBuilder.ToString();
        //     return txt;
        // }

        // static private string ToLuaData(string cellType, string cellValue)
        // {
        //     switch (cellType)
        //     {
        //         case Xlsx2LuaConfig.Number:
        //             cellValue = LuaTemplate.OBJ.Format(cellValue);
        //             break;
        //         case Xlsx2LuaConfig.String:
        //             cellValue = LuaTemplate.STR.Format(cellValue);
        //             break;
        //         case Xlsx2LuaConfig.Table:

        //             break;
        //     }
        //     return cellValue;
        // }

        // /// <summary>
        // /// row1: 该表的描述说明
        // /// row2: key、type、value、desc
        // /// </summary>
        // /// <param name="nameSpace"></param>
        // /// <param name="tableName"></param>
        // /// <param name="sheet"></param>
        // /// <returns></returns>
        // static private string BuildLuaObj(string nameSpace, string tableName, DataTable sheet)
        // {
        //     int columnsNum = sheet.Columns.Count;
        //     int rowsNum = sheet.Rows.Count;
        //     if (rowsNum < 2) throw new Exception("表内容格式不正确！！！\n" + "row1: 该表的描述说明\n" + "row2: key(参数名)、type(参数类型)、value(参数值)、desc(参数说明)");
        //     LuaBuilder luaBuilder = new LuaBuilder();
        //     for (int i = 2; i < rowsNum; ++i)
        //     {
        //         DataRow row = sheet.Rows[i];
        //         string argName = row[0].ToString();
        //         string argType = row[1].ToString();
        //         string argValue = row[2].ToString();
        //         string argDesc = row[3].ToString();
        //         luaBuilder.AddDesc(argDesc);
        //         string fieldValue = ToLuaData(argType, argValue);
        //         luaBuilder.AddObjField(argName, fieldValue);
        //     }
        //     return luaBuilder.ToLocalTbl(tableName);
        // }

        // /// <summary>
        // /// row1: 该表的描述说明
        // /// row2: key、value、desc
        // /// </summary>
        // /// <param name="nameSpace"></param>
        // /// <param name="tableName"></param>
        // /// <param name="sheet"></param>
        // /// <returns></returns>
        // static private string BuildLuaEnum(string nameSpace, string tableName, DataTable sheet)
        // {
        //     int columnsNum = sheet.Columns.Count;
        //     int rowsNum = sheet.Rows.Count;
        //     if (rowsNum < 2) throw new Exception("表内容格式不正确！！！\n" + "row1: 该表的描述说明\n" + "row2: key(参数名)、value(参数值)、desc(参数说明)");
        //     LuaBuilder luaBuilder = new LuaBuilder();
        //     int enumIndex = -1;
        //     for (int i = 2; i < rowsNum; ++i)
        //     {
        //         DataRow row = sheet.Rows[i];
        //         string argName = row[0].ToString();
        //         enumIndex = row.IsNull(2) ? ++enumIndex : int.Parse(row[2].ToString());
        //         string argValue = enumIndex.ToString();
        //         string argDesc = row[3].ToString();
        //         luaBuilder.AddDesc(argDesc);
        //         string fieldValue = ToLuaData(Xlsx2LuaConfig.Number, argValue);
        //         luaBuilder.AddObjField(argName, fieldValue);
        //     }
        //     return luaBuilder.ToLocalTbl(tableName);
        // }

        // /// <summary>
        // /// row1: 该表的描述说明\n
        // /// row2: 参数说明
        // /// row3: 参数类型
        // /// row3: 参数名
        // /// </summary>
        // /// <param name="nameSpace"></param>
        // /// <param name="tableName"></param>
        // /// <param name="sheet"></param>
        // /// <returns></returns>
        // static private string BuildLuaTable(string nameSpace, string tableName, DataTable sheet)
        // {
        //     int columnsNum = sheet.Columns.Count;
        //     int rowsNum = sheet.Rows.Count;
        //     if (rowsNum < 4) throw new Exception("表内容格式不正确！！！\n" + "row1: 该表的描述说明\n" + "row2: 参数说明\n" + "row3: 参数类型\n" + "row3: 参数名\n");

        //     string sheetName = sheet.TableName;
        //     LuaBuilder luaBuilder = new LuaBuilder();
        //     DataRow tblDesc = sheet.Rows[0];
        //     luaBuilder.AddDesc(tblDesc[0].ToString());

        //     DataRow fieldDesc;
        //     DataRow fieldTypes;
        //     DataRow fieldNames;
        //     List<string> descList = new List<string>();
        //     List<string> typeList = new List<string>();
        //     List<string> nameList = new List<string>();

        //     LuaBuilder objBuilder = new LuaBuilder();
        //     LuaBuilder tblBuilder = new LuaBuilder();

        //     fieldDesc = sheet.Rows[1];
        //     fieldTypes = sheet.Rows[2];
        //     fieldNames = sheet.Rows[3];
        //     for (int i = 0; i < columnsNum; ++i)
        //     {
        //         string desc = fieldDesc.IsNull(i) ? string.Empty : fieldDesc[i].ToString();
        //         string typeName = fieldTypes.IsNull(i) ? throw new NoNullAllowedException("sheet 缺少字段类型: {0}".Format(sheet.TableName)) : fieldTypes[i].ToString();
        //         string fieldName = fieldNames.IsNull(i) ? throw new NoNullAllowedException("sheet 缺少字段名: {0}".Format(sheet.TableName)) : fieldNames[i].ToString();
        //         descList.Add(desc);
        //         typeList.Add(typeName);
        //         nameList.Add(fieldName);
        //         objBuilder.AddObjField(fieldName, i.ToString());
        //     }

        //     luaBuilder.AddSubContent(objBuilder.ToLocalTbl(tableName));

        //     for (int i = 4, enumIndex = 0; i < rowsNum; ++i, ++enumIndex)
        //     {
        //         DataRow row = sheet.Rows[i];
        //         string key = row[0].ToString();

        //         List<string> rowValues = new List<string>();
        //         for (int j = 1; j < columnsNum; ++j)
        //         {
        //             string cellType = typeList[j];
        //             string cellValue = row[j].ToString();
        //             cellValue = ToLuaData(cellType, cellValue);
        //             rowValues.Add(cellValue);
        //         }
        //         string tblItemKey = row[0].ToString();
        //         string tblItemValue = LuaBuilder.ToLuaTable(rowValues.ToArray());
        //         tblBuilder.AddListItem(tblItemKey, tblItemValue);
        //         break;
        //     }

        //     string txt = luaBuilder.ToString();
        //     return txt;
        // }
    }
}
