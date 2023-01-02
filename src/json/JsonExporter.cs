using System.Security.Cryptography.X509Certificates;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

using ExcelDataReader;

namespace GFramework.Xlsx
{
    public class JsonExporter
    {
        private GLogger logger = new GLogger("JsonExporter");

        public Dictionary<string, Dictionary<string, XlsxDataModel>> objSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
        public Dictionary<string, Dictionary<string, XlsxDataModel>> enumSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
        public Dictionary<string, Dictionary<string, XlsxDataModel>> tblSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
        public Dictionary<string, LinkData> linkMap = new Dictionary<string, LinkData>();

        public XlsxConfig xlsxCfg = null;
        public JsonConfig jsonCfg = null;

        public JsonExporter( Dictionary<string, Dictionary<string, XlsxDataModel>> objSheetMap, Dictionary<string, Dictionary<string, XlsxDataModel>> enumSheetMap, Dictionary<string, Dictionary<string, XlsxDataModel>> tblSheetMap, Dictionary<string, LinkData> linkMap, XlsxConfig xlsxCfg, JsonConfig jsonCfg)
        {
            this.objSheetMap = objSheetMap;
            this.enumSheetMap = enumSheetMap;
            this.tblSheetMap = tblSheetMap;
            this.linkMap = linkMap;
            this.xlsxCfg = xlsxCfg;
            this.jsonCfg = jsonCfg;
        }

        public void ExportToJson()
        {
            string luaPath = xlsxCfg.LuaConfig.ExportTo;
            if (!Directory.Exists(luaPath))
                Directory.CreateDirectory(luaPath);
            logger.P("导出路径：{0}".Format(luaPath));

            this.OutputJsonObj();
            this.OutputJsonEnum();
            this.OutputJsonTbl();
            this.LinkInlineTbl();
            this.WriteAllXlsxDataToLua();
        }

        //生成lua 对象
        private void OutputJsonObj()
        {
            foreach (var nameSpace in this.objSheetMap.Keys)
            {
                //根据命名空间创建目录
                string exportTo = this.jsonCfg.ExportTo;
                string dirPath = Path.Combine(exportTo, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var objMap = this.objSheetMap[nameSpace];
                //导出sheet
                foreach (var tblName in objMap.Keys)
                {
                    var dataModel = objMap[tblName];
                    var objDataList = dataModel.objDataList;
                    JsonBuilder jsonBuilder = new JsonBuilder();
                    for (int i = 0; i < objDataList.Count; ++i)
                    {
                        var objData = objDataList[i];
                        string fieldName = objData.fieldName;     // 字段名称
                        string fieldType = objData.fieldType;     // 字段类型
                        string fieldValue = objData.fieldValue;    // 字段值
                        string fieldDesc = objData.fieldDesc;     // 字段表述

                        // 根据数据类型生成lua字段
                        fieldValue = ToJsonData(fieldType, fieldValue);
                        // 添加字段
                        jsonBuilder.AddObjField(fieldName, fieldValue);
                    }

                    //lua文件名格式
                    string jsonPath = Path.Combine(curDir.FullName, tblName + this.jsonCfg.Externion).PathFormat();
                    dataModel.export = jsonPath;

                    //lua导出包
                    string packageName = this.xlsxCfg.DataTableFormat.Format(nameSpace.Upper(), tblName.Upper());
                    string luaData = JsonBuilder.ToLocalTbl(packageName, jsonBuilder.ToString());
                    // 添加表说明
                    string luaTxt = JsonBuilder.Package(luaData);

                    string exportRootDir = this.jsonCfg.ExportTo;
                    dataModel.txt = luaTxt;
                }
            }
        }

        //生成lua 枚举
        private void OutputJsonEnum()
        {
            foreach (var nameSpace in this.enumSheetMap.Keys)
            {
                //根据命名空间创建目录
                string exportTo = this.jsonCfg.ExportTo;
                string dirPath = Path.Combine(exportTo, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var enumMap = this.enumSheetMap[nameSpace];
                //导出sheet
                foreach (var tblName in enumMap.Keys)
                {
                    var dataModel = enumMap[tblName];
                    var objDataList = dataModel.objDataList;
                    JsonBuilder jsonBuilder = new JsonBuilder();
                    for (int i = 0; i < objDataList.Count; ++i)
                    {
                        var objData = objDataList[i];
                        string fieldName = objData.fieldName;     // 字段名称
                        string fieldType = objData.fieldType;     // 字段类型
                        string fieldValue = objData.fieldValue;    // 字段值
                        string fieldDesc = objData.fieldDesc;     // 字段表述

                        // 根据数据类型生成lua字段
                        fieldValue = ToJsonData(fieldType, fieldValue);
                        // 添加字段
                        jsonBuilder.AddObjField(fieldName, fieldValue);
                    }

                    //lua文件名格式
                    string luaPath = Path.Combine(curDir.FullName, tblName + this.jsonCfg.Externion).PathFormat();
                    dataModel.export = luaPath;

                    //lua导出包
                    string packageName = this.xlsxCfg.DataTableFormat.Format(nameSpace.Upper(), tblName.Upper());
                    string luaData = JsonBuilder.ToLocalTbl(packageName, jsonBuilder.ToString());
                    // 添加表说明
                    string luaTxt = JsonBuilder.Package(luaData);

                    string exportRootDir = this.jsonCfg.ExportTo;
                    dataModel.txt = luaTxt;
                }
            }
        }

        //生成lua 数据表
        private void OutputJsonTbl()
        {
            foreach (var nameSpace in this.tblSheetMap.Keys)
            {
                //根据命名空间创建目录
                string exportTo = this.jsonCfg.ExportTo;
                string dirPath = Path.Combine(exportTo, nameSpace).PathFormat();
                var exportDir = Directory.CreateDirectory(dirPath);

                var tblMap = this.tblSheetMap[nameSpace];
                //导出sheet
                foreach (var tblName in tblMap.Keys)
                {
                    var dataModel = tblMap[tblName];
                    var objDataList = dataModel.objDataList;
                    string packageName = this.xlsxCfg.DataTableFormat.Format(nameSpace.Upper(), tblName.Upper());

                    JsonBuilder objBuilder = new JsonBuilder();
                    int xid = -1;
                    for (int i = 0; i < objDataList.Count; ++i)
                    {
                        var objData = objDataList[i];
                        string fieldName = objData.fieldName;
                        string fieldType = objData.fieldType;
                        string fieldDesc = objData.fieldDesc;

                        if (fieldType == this.xlsxCfg.XlsxTypes.Xid)
                        {
                            xid = i;
                        }
                        if (objData.group != null) continue;
                        objBuilder.AddObjField(fieldName, i.ToString());
                    }

                    //添加表对象
                    string dataModelName = this.xlsxCfg.DataTableObjectFormat.Format(tblName.Upper());
                    string luaTblObj = JsonBuilder.ToLocalTbl(dataModelName, objBuilder.ToString());

                    // 分组声明文件

                    //生成数据项
                    var tblDataMap = dataModel.tblDataMap;
                    JsonBuilder dataBuilder = new JsonBuilder();
                    foreach (var index in tblDataMap.Keys)
                    {
                        JsonBuilder itemBuilder = new JsonBuilder();

                        var tblItems = dataModel.tblDataMap[index].data;
                        foreach (var fn in tblItems.Keys)
                        {
                            var fieldItem = tblItems[fn];
                            // 字段名
                            string filedName = fieldItem.fieldName;
                            // 字段类型
                            string fieldType = fieldItem.fieldType;
                            // 字段值
                            string fieldValue = fieldItem.fieldValue;
                            if (fieldType == this.xlsxCfg.XlsxTypes.Xid)
                                continue;

                            fieldItem.fieldValue = ToJsonData(fieldType, fieldValue);
                            // 添加字段
                            itemBuilder.AddObjField(filedName, fieldItem.fieldValue);
                        }

                        string dataItem = JsonTemplate.OBJ.Format(itemBuilder.ToString());
                        // 添加数据项
                        dataBuilder.AddObjField(index, dataItem);
                    }

                    // lua 文件名
                    string lua = Path.Combine(exportDir.FullName, tblName + jsonCfg.Externion).PathFormat();
                    dataModel.export = lua;

                    //lua导出包
                    string luaTblData = JsonBuilder.ToLocalTbl(packageName, dataBuilder.ToString());
                    string luaTxt = JsonBuilder.Package(luaTblData);

                    //lua文件描述
                    string exportRootDir = this.jsonCfg.ExportTo;
                    dataModel.txt = luaTxt;
                }
            }
        }

        private string ToJsonData(string fieldType, string fieldValue)
        {
            fieldType = fieldType.Trim();
            if (string.IsNullOrEmpty(fieldValue))
            {
                fieldValue = JsonTemplate.NIL;
                return fieldValue;
            }
            XlsxTypes xlsxTypes = this.xlsxCfg.XlsxTypes;
            if (fieldType == xlsxTypes.Number) fieldValue = JsonTemplate.NUM.Format(fieldValue);
            else if (fieldType == xlsxTypes.String) fieldValue = JsonTemplate.STR.Format(fieldValue);
            else if (fieldType == xlsxTypes.ListNumber)
            {
                var nums = fieldValue.Split(',');
                for (int i = 0; i < nums.Length; ++i)
                {
                    string v = nums[i].ToString().Trim();
                    if (!v.IsDigit())
                    {
                        logger.E("数值数组包含非法字符！！！");
                        continue;
                    }
                }
                fieldValue = JsonTemplate.LIST.Format(string.Join(',', nums));
            }
            else if (fieldType == xlsxTypes.ListString)
            {
                var strs = fieldValue.Split(',');
                for (int i = 0; i < strs.Length; ++i)
                {
                    string v = strs[i].ToString().Trim();
                    if (!v.IsAlpha())
                    {
                        logger.E("字符串数组包含非法字符！！！");
                        continue;
                    }
                    strs[i] = JsonTemplate.STR.Format(v);
                }
                fieldValue = JsonTemplate.LIST.Format(string.Join(',', strs));
            }
            else if (fieldType == xlsxTypes.Xid) { }
            else if (Regex.IsMatch(fieldType, xlsxTypes.InlineTable))
            {
                var match = Regex.Match(fieldType, this.xlsxCfg.XlsxTypes.InlineTable);
                string inlineTblNameSpace = match.Groups[1].Value;
                string inlineTblName = match.Groups[2].Value;
                fieldValue = XlsxExporter.InlineTblRegexFormat.Format(inlineTblNameSpace, inlineTblName, fieldValue);
            }
            else
            {
                logger.E("配置表数据类型错误！！！");
            }
            return fieldValue;
        }

        /// <summary>
        /// 链接内嵌表
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
                sourceDataModel.txt = sourceDataModel.txt.Replace(inlineTblIndex, inlineTxt);
            }
        }

        private string DepLinkInlineTbl(LinkData rootData, LinkData linkData)
        {
            var sourceTbl = linkData.sourceTbl;
            string sourceNameSpace = sourceTbl.nameSpace;
            string sourceTblName = sourceTbl.tblName;
            string sourceDataIndex = sourceTbl.dataIndex;
            var sourceTblModel = this.tblSheetMap[sourceNameSpace][sourceTblName];
            string luaTxt = sourceTblModel.txt;

            var inlineTbl = linkData.inlineTbl;
            string inlineNameSpace = inlineTbl.nameSpace;
            string inlineTblName = inlineTbl.tblName;
            string inlineDataIndex = inlineTbl.dataIndex;
            var inlineDataIndexList = inlineDataIndex.Split(xlsxCfg.XlsxTypes.ListSeparator);

            string inlineTblIndexDataTxt = inlineDataIndex;
            JsonBuilder inlineTblBuilder = new JsonBuilder();
            foreach (var index in inlineDataIndexList)
            {
                var indexData = this.tblSheetMap[inlineNameSpace][inlineTblName].tblDataMap[index];
                var indexDataTxt = JsonBuilder.ToTbl(indexData.data);

                bool hasInline = Regex.IsMatch(indexDataTxt, XlsxExporter.InlineTblRegex);
                if (hasInline)
                {
                    var matchs = Regex.Matches(indexDataTxt, XlsxExporter.InlineTblRegex);
                    foreach (Match nextMatch in matchs)
                    {
                        var nextInlineTblIndex = nextMatch.Groups[1].Value;
                        var matchTblIndex3 = Regex.Match(nextInlineTblIndex, XlsxExporter.InlineTblBelongRegex);
                        var nameSpace = matchTblIndex3.Groups[1].Value;
                        var tblName = matchTblIndex3.Groups[2].Value;
                        var dataIndex = matchTblIndex3.Groups[3].Value;

                        var rootTbl = rootData.sourceTbl;
                        //循环嵌套
                        if (nameSpace == rootTbl.nameSpace && tblName == rootTbl.tblName)
                        {
                            var tempTblData = new XlsxTblData(indexData.data);
                            for (int i = 0; i < tempTblData.data.Count; i++)
                            {
                                var tmp = tempTblData.data.ElementAt(i);
                                string fieldType = tmp.Value.fieldType;
                                if (fieldType == XlsxExporter.InlineTblBelongRegexFormat.Format(nameSpace, tblName) || fieldType == this.xlsxCfg.XlsxTypes.Xid)
                                    tempTblData.data.Remove(tmp.Key);
                            }
                            indexDataTxt = JsonBuilder.ToTbl(tempTblData.data);
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
                inlineTblBuilder.AddObjField(index, indexDataTxt);
            }
            return JsonTemplate.OBJ.Format(inlineTblBuilder.ToString());
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

        private void WriteToLua(XlsxDataModel dataModel)
        {
            string xlsx = dataModel.xlsx;
            string lua = dataModel.export;
            string luaTxt = dataModel.txt;
            using (FileStream fs = new FileStream(lua, FileMode.OpenOrCreate, FileAccess.Write))
            {
                fs.SetLength(0);
                byte[] luaData = Encoding.UTF8.GetBytes(luaTxt);
                fs.Write(luaData);
            }
            logger.P("导出完成{0}...".Format(dataModel.export));
        }
    }
}