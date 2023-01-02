using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;
using GFramework;
using Newtonsoft.Json.Linq;

namespace GFramework.Xlsx
{
    public class XlsxExporter
    {
        private GLogger logger = new GLogger("LuaExporter");

        public const string InlineTblRegexFormat = "{0}::{1}::{2}";
        public const string InlineTblBelongRegexFormat = "{0}::{1}";
        public const string InlineTblRegex = "([.*?^[a-zA-Z_][a-zA-Z0-9_]*\\s*::\\s*[a-zA-Z_][a-zA-Z0-9_]*\\s*::\\s*[\\s+a-zA-Z0-9_,]+)";
        public const string InlineTblBelongRegex = "^([a-zA-Z_][a-zA-Z0-9_]*)\\s*::\\s*([a-zA-Z_][a-zA-Z0-9_]*)\\s*::\\s*([\\s+a-zA-Z0-9_,]+)";

        private JObject declareJson = null;

        private XlsxConfig xlsxCfg = null;

        public void ExecuteExport()
        {
            //加载配置表
            var xlsxCfg = LoadConfig();
            this.xlsxCfg = xlsxCfg;

            //检查xlsx目录
            string sourcePath = Path.GetFullPath(xlsxCfg.Xlsx).PathFormat();
            if (!Directory.Exists(sourcePath))
                throw new DirectoryNotFoundException(sourcePath);
            logger.P("xlsx路径：{0}".Format(sourcePath));

            if (!Regex.IsMatch(xlsxCfg.ExportCmd, "lua|json|cs"))
                logger.E("不支持该导出命令：{0}".Format(xlsxCfg.ExportCmd));
            var xlsxs = FindAllXlsx();
            string[] exportCmd = xlsxCfg.ExportCmd.Split('|');
            foreach (string cmd in exportCmd)
            {
                switch (cmd)
                {
                    case "lua":
                        logger.P("执行导出lua...");
                        var (objMap1, enumMap1, tblMap1, linkMap1) = OutputXlsxDataModels(xlsxs);
                        // var objMap1 = MemoryManager.DeepClone(objSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var enumMap1 = MemoryManager.DeepClone(enumSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var tblMap1 = MemoryManager.DeepClone(tblSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var linkMap1 = MemoryManager.DeepClone(linkMap) as Dictionary<string, LinkData>;

                        var luaCfg = xlsxCfg.LuaConfig;
                        luaCfg.ExportTo = Path.GetFullPath(luaCfg.ExportTo);

                        string declarePath = Path.GetFullPath(luaCfg.DeclareJson);
                        if (!File.Exists(declarePath)) File.Create(declarePath);
                        this.declareJson = new JsonStream(declarePath).Read<JObject>();

                        LuaExporter lua = new LuaExporter(objMap1, enumMap1, tblMap1, linkMap1, xlsxCfg, luaCfg, declareJson);
                        lua.ExportToLua();
                        break;
                    case "json":
                        logger.P("执行导出json...");
                        var (objMap2, enumMap2, tblMap2, linkMap2) = OutputXlsxDataModels(xlsxs);
                        // var objMap2 = MemoryManager.DeepClone(objSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var enumMap2 = MemoryManager.DeepClone(enumSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var tblMap2 = MemoryManager.DeepClone(tblSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var linkMap2 = MemoryManager.DeepClone(linkMap) as Dictionary<string, LinkData>;

                        var jsonCfg = xlsxCfg.JsonConfig;
                        jsonCfg.ExportTo = Path.GetFullPath(jsonCfg.ExportTo);
                        JsonExporter json = new JsonExporter(objMap2, enumMap2, tblMap2, linkMap2, xlsxCfg, jsonCfg);
                        json.ExportToJson();
                        break;
                    case "cs":
                        logger.P("执行导出cs...");
                        var (objMap3, enumMap3, tblMap3, linkMap3) = OutputXlsxDataModels(xlsxs);
                        // var objMap2 = MemoryManager.DeepClone(objSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var enumMap2 = MemoryManager.DeepClone(enumSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var tblMap2 = MemoryManager.DeepClone(tblSheetMap) as Dictionary<string, Dictionary<string, XlsxDataModel>>;
                        // var linkMap2 = MemoryManager.DeepClone(linkMap) as Dictionary<string, LinkData>;

                        var csCfg = xlsxCfg.CSConfig;
                        csCfg.ExportTo = Path.GetFullPath(csCfg.ExportTo);
                        CSExporter cs = new CSExporter(objMap3, enumMap3, tblMap3, xlsxCfg, csCfg);
                        cs.ExportToCS();
                        break;
                    default:
                        break;
                }
            }
        }

        public XlsxConfig LoadConfig()
        {
            string curDir = Environment.CurrentDirectory;
            string cfgName = "xlsx.config.json";
            string cfgPath = Path.Combine(curDir, cfgName).Format();
            JsonStream js = new JsonStream(cfgPath);
            var cfg = js.Read<XlsxConfig>();
            return cfg;
        }

        /// <summary>
        /// 获取所有xlsx文件
        /// </summary>
        private List<string> FindAllXlsx()
        {
            List<string> files = new List<string>();
            var xlsxs = FileUtil.FindFilesByTypeWithDep(this.xlsxCfg.Xlsx, ".xlsx");

            logger.P("xlsx 列表：");
            for (int i = 0; i < xlsxs.Count; ++i)
            {
                string xlsx = xlsxs[i];
                logger.P("{0}：{1}".Format(i, xlsx));
            }
            logger.P("共{0}个xlsx文件".Format(xlsxs.Count));

            return xlsxs;
        }

        /// <summary>
        /// 获取所有类型的sheet
        /// </summary>
        private (Dictionary<string, Dictionary<string, XlsxDataModel>>,
        Dictionary<string, Dictionary<string, XlsxDataModel>>,
        Dictionary<string, Dictionary<string, XlsxDataModel>>,
        Dictionary<string, LinkData>)
        OutputXlsxDataModels(List<string> xlsxs)
        {
            var objSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
            var enumSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
            var tblSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
            var linkMap = new Dictionary<string, LinkData>();
            foreach (var xlsx in xlsxs)
            {
                //忽略不需要导出文件
                string xlsxFileName = xlsx.GetCurFileName();
                if (Regex.IsMatch(xlsxFileName, this.xlsxCfg.IgnoreXlsxRegex))
                    continue;


                //获取所有sheet
                Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (FileStream fs = new FileStream(xlsx, FileMode.Open))
                {
                    //匹配命名空间
                    var matchNameSpace = Regex.Match(xlsxFileName, this.xlsxCfg.NameSpaceRegex);
                    string nameSpace = matchNameSpace.Groups[1].Value;
                    //默认全局命名空间
                    if (string.IsNullOrEmpty(nameSpace))
                        nameSpace = this.xlsxCfg.LuaDefaultNameSpace;

                    var excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    var sheetSet = excelDataReader.AsDataSet().Tables;
                    foreach (DataTable sheet in sheetSet)
                    {
                        string sheetName = sheet.TableName;
                        string tblName = sheetName.ToVaildName();

                        //忽略不需要导出的表
                        if (Regex.IsMatch(sheetName, this.xlsxCfg.IgnoreSheetRegex))
                            continue;

                        string luaType = string.Empty;
                        string desc = string.Empty;
                        int rowStart = -1;
                        int colStart = -1;
                        int rowsNum = sheet.Rows.Count;
                        int colsNum = sheet.Columns.Count;
                        for (int r = 0; r < rowsNum; ++r)
                        {
                            bool find = false;
                            for (int c = 0; c < colsNum; ++c)
                            {
                                var cell = sheet.Rows[r][c].ToString();
                                if (string.IsNullOrEmpty(cell)) continue;
                                if (Regex.IsMatch(cell, this.xlsxCfg.XlsxTypeRegex))
                                {
                                    var xlsxTypeMatch = Regex.Match(cell, this.xlsxCfg.XlsxTypeRegex);
                                    luaType = xlsxTypeMatch.Groups[1].Value;
                                    desc = xlsxTypeMatch.Groups[2].Value;
                                    rowStart = r;
                                    colStart = c;
                                    find = true;
                                    break;
                                }
                            }
                            if (find) break;
                        }

                        XlsxDataModel dataModel = new XlsxDataModel()
                        {
                            xlsx = xlsx,
                            nameSpace = nameSpace,
                            tblName = tblName,
                            xlsxType = luaType,
                            desc = desc,
                        };

                        switch (luaType)
                        {
                            case "OBJ":
                                if (!objSheetMap.ContainsKey(nameSpace)) objSheetMap.Add(nameSpace, new Dictionary<string, XlsxDataModel>());
                                if (objSheetMap[nameSpace].ContainsKey(tblName))
                                {
                                    logger.E("导出表名冲突：{0}/{1}".Format(xlsx, tblName));
                                    continue;
                                }
                                DealWithXlsxObj(sheet, rowStart, colStart, ref dataModel);
                                objSheetMap[nameSpace][tblName] = dataModel;
                                break;
                            case "ENUM":
                                if (!enumSheetMap.ContainsKey(nameSpace)) enumSheetMap.Add(nameSpace, new Dictionary<string, XlsxDataModel>());
                                if (enumSheetMap[nameSpace].ContainsKey(tblName))
                                {
                                    logger.E("导出表名冲突：{0}/{1}".Format(xlsx, tblName));
                                    continue;
                                }
                                DealWithXlsxEnum(sheet, rowStart, colStart, ref dataModel);
                                enumSheetMap[nameSpace][tblName] = dataModel;
                                break;
                            case "TBL":
                                if (!tblSheetMap.ContainsKey(nameSpace)) tblSheetMap.Add(nameSpace, new Dictionary<string, XlsxDataModel>());
                                if (tblSheetMap[nameSpace].ContainsKey(tblName))
                                {
                                    logger.E("导出表名冲突：{0}/{1}".Format(xlsx, tblName));
                                    continue;
                                }
                                DealWithXlsxTbl(sheet, rowStart, colStart, ref dataModel, ref linkMap);
                                tblSheetMap[nameSpace][tblName] = dataModel;
                                break;
                            default:
                                logger.E("不支持导出该lua数据类型：{0}".Format(luaType));
                                break;
                        }
                    }
                }
            }
            return (objSheetMap, enumSheetMap, tblSheetMap, linkMap);
        }

        public void DealWithXlsxObj(DataTable sheet, int rowStart, int colStart, ref XlsxDataModel dataModel)
        {
            int columnsNum = sheet.Columns.Count;
            int rowsNum = sheet.Rows.Count;

            for (int i = rowStart + 2; i < rowsNum; ++i)
            {
                DataRow row = sheet.Rows[i];
                string fieldName = row[0].ToString();     // 字段名称
                string fieldType = row[1].ToString();     // 字段类型
                string fieldValue = row[2].ToString();    // 字段值
                string fieldDesc = row[3].ToString();     // 字段表述
                // 格式检查
                if (string.IsNullOrEmpty(fieldName))
                {
                    logger.E("缺少字段名称：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }
                if (string.IsNullOrEmpty(fieldType))
                {
                    logger.E("缺少字段类型：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }
                if (string.IsNullOrEmpty(fieldValue))
                {
                    logger.E("缺少字段值：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }
                var objItem = new XlsxTblItemData(fieldType, fieldName, fieldValue, fieldDesc);
                dataModel.objDataList.Add(objItem);
            }
        }

        private void DealWithXlsxEnum(DataTable sheet, int rowStart, int colStart, ref XlsxDataModel dataModel)
        {
            int columnsNum = sheet.Columns.Count;
            int rowsNum = sheet.Rows.Count;

            for (int i = rowStart + 2, enumIndex = 0; i < rowsNum; ++i, ++enumIndex)
            {
                DataRow row = sheet.Rows[i];
                string fieldName = row[0].ToString();     // 字段名称
                string fieldValue = row[1].ToString();    // 字段值
                string fieldDesc = row[2].ToString();     // 字段表述
                // 格式检查
                if (string.IsNullOrEmpty(fieldName))
                {
                    logger.E("缺少枚举名：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }

                if (string.IsNullOrEmpty(fieldValue))
                {
                    logger.E("缺少枚举值：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }

                var objItem = new XlsxTblItemData(this.xlsxCfg.XlsxTypes.Number, fieldName, fieldValue, fieldDesc);
                dataModel.objDataList.Add(objItem);
            }
        }

        private void DealWithXlsxTbl(DataTable sheet, int rowStart, int colStart, ref XlsxDataModel dataModel, ref Dictionary<string, LinkData> linkMap)
        {
            int columnsNum = sheet.Columns.Count;
            int rowsNum = sheet.Rows.Count;

            //生成表对象
            DataRow fieldNames = sheet.Rows[rowStart + 1];
            DataRow fieldTypes = sheet.Rows[rowStart + 2];
            DataRow fieldDescs = sheet.Rows[rowStart + 3];

            int xid = -1;
            for (int c = colStart, enumIndex = 0; c < columnsNum; ++c, ++enumIndex)
            {
                string fieldName = fieldNames[c].ToString();
                string fieldType = fieldTypes[c].ToString();
                string fieldDesc = fieldDescs[c].ToString();

                // 格式检查
                if (string.IsNullOrEmpty(fieldType))
                {
                    logger.E("缺少字段类型：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }

                if (string.IsNullOrEmpty(fieldName))
                {
                    logger.E("缺少字段名称：{0}/{1}".Format(dataModel.xlsx, dataModel.tblName));
                    continue;
                }

                if (fieldType == this.xlsxCfg.XlsxTypes.Xid)
                {
                    xid = c;
                }

                string fieldValue = enumIndex.ToString();
                var objItem = new XlsxTblItemData(fieldType, fieldName, fieldValue, fieldDesc);
                dataModel.objDataList.Add(objItem);
            }

            var objDataList = dataModel.objDataList;
            for (int r = rowStart + 4, enumIndex = 0; r < rowsNum; ++r, ++enumIndex)
            {
                DataRow row = sheet.Rows[r];

                // 表索引
                string key = xid < 0 ? enumIndex.ToString() : row[xid].ToString();
                dataModel.tblDataMap[key] = new XlsxTblData();

                for (int c = colStart, fieldIndex = 0; c < columnsNum; ++c, ++fieldIndex)
                {
                    if (c == xid) continue;
                    // 字段名
                    string fieldName = objDataList[fieldIndex].fieldName;
                    // 字段类型
                    string fieldType = objDataList[fieldIndex].fieldType;
                    // 字段值
                    string fieldValue = row[c].ToString();

                    //内嵌表
                    bool isInline = Regex.IsMatch(fieldType, this.xlsxCfg.XlsxTypes.InlineTable);
                    if (isInline)
                    {
                        var match = Regex.Match(fieldType, this.xlsxCfg.XlsxTypes.InlineTable);
                        string inlineTblNameSpace = match.Groups[1].Value;
                        string inlineTblName = match.Groups[2].Value;
                        LinkData linkData = new LinkData(
                            new SourceTbl(dataModel.nameSpace, dataModel.tblName, key),
                            new InlineTbl(inlineTblNameSpace, inlineTblName, fieldValue)
                        );
                        linkMap[linkData.sourceTbl.ToString()] = linkData;
                    }

                    var tblItem = new XlsxTblItemData(fieldType, fieldName, fieldValue);
                    dataModel.tblDataMap[key].data[fieldName] = tblItem;
                }
            }
        }
    }
}