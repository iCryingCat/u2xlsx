using System;
using System.Text;
using System.Text.RegularExpressions;

namespace GFramework.Xlsx
{
    public class CSExporter
    {
        private GLogger logger = new GLogger("CSExporter");

        public Dictionary<string, Dictionary<string, XlsxDataModel>> objSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
        public Dictionary<string, Dictionary<string, XlsxDataModel>> enumSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();
        public Dictionary<string, Dictionary<string, XlsxDataModel>> tblSheetMap = new Dictionary<string, Dictionary<string, XlsxDataModel>>();

        public XlsxConfig xlsxCfg = null;
        public CSConfig csCfg = null;

        public CSExporter(Dictionary<string, Dictionary<string, XlsxDataModel>> objSheetMap, Dictionary<string, Dictionary<string, XlsxDataModel>> enumSheetMap, Dictionary<string, Dictionary<string, XlsxDataModel>> tblSheetMap, XlsxConfig xlsxCfg, CSConfig csCfg)
        {
            this.objSheetMap = objSheetMap;
            this.enumSheetMap = enumSheetMap;
            this.tblSheetMap = tblSheetMap;
            this.xlsxCfg = xlsxCfg;
            this.csCfg = csCfg;
        }

        public void ExportToCS()
        {
            string csPath = csCfg.ExportTo;
            if (!Directory.Exists(csPath))
                Directory.CreateDirectory(csPath);
            logger.P("导出路径：{0}".Format(csPath));

            this.OutputCSObj();
            this.OutputCSEnum();
            this.OutputCSTbl();
            this.WriteAllXlsxDataToCS();
        }

        //生成cs 对象
        private void OutputCSObj()
        {
            foreach (var nameSpace in this.objSheetMap.Keys)
            {
                //根据命名空间创建目录
                string exportTo = this.csCfg.ExportTo;
                string dirPath = Path.Combine(exportTo, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var objMap = this.objSheetMap[nameSpace];
                //导出sheet
                foreach (var tblName in objMap.Keys)
                {
                    var dataModel = objMap[tblName];
                    var objDataList = dataModel.objDataList;
                    CSBuilder classBuilder = new CSBuilder();
                    CSBuilder usingBuilder = new CSBuilder();
                    for (int i = 0; i < objDataList.Count; ++i)
                    {
                        var objData = objDataList[i];
                        string fieldName = objData.fieldName;     // 字段名称
                        string fieldType = objData.fieldType;     // 字段类型
                        string fieldDesc = objData.fieldDesc;     // 字段表述

                        // 根据数据类型生成cs字段
                        fieldType = ToCSData(fieldType, ref usingBuilder);
                        // 添加字段
                        classBuilder.AddDesc(fieldDesc);
                        classBuilder.AddField(fieldType, fieldName);
                    }

                    //cs文件名格式
                    string csPath = Path.Combine(curDir.FullName, tblName + this.csCfg.Externion).PathFormat();
                    dataModel.export = csPath;

                    //cs导出包
                    string classBody = CSBuilder.PackageClass(tblName.Upper(), classBuilder.ToString());
                    string packageName = this.csCfg.PackageFormat.Format(nameSpace.Upper());
                    string classDesc = CSTemplate.DESC.Format(dataModel.desc);
                    string nameSpaceBody = CSBuilder.PackageNameSpace(packageName, classDesc + classBody);
                    string usingBody = usingBuilder.ToString();
                    dataModel.txt = usingBody + nameSpaceBody;
                }
            }
        }

        //生成cs 枚举
        private void OutputCSEnum()
        {
            foreach (var nameSpace in this.enumSheetMap.Keys)
            {
                //根据命名空间创建目录
                string exportTo = this.csCfg.ExportTo;
                string dirPath = Path.Combine(exportTo, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var enumMap = this.enumSheetMap[nameSpace];
                //导出sheet
                foreach (var tblName in enumMap.Keys)
                {
                    var dataModel = enumMap[tblName];
                    var objDataList = dataModel.objDataList;
                    CSBuilder classBuilder = new CSBuilder();
                    CSBuilder usingBuilder = new CSBuilder();
                    for (int i = 0; i < objDataList.Count; ++i)
                    {
                        var objData = objDataList[i];
                        string fieldDesc = objData.fieldDesc;     // 字段表述
                        string fieldName = objData.fieldName;     // 字段名称
                        string fieldValue = objData.fieldValue; // 字段值

                        // 添加字段
                        classBuilder.AddDesc(fieldDesc);
                        classBuilder.AddEnum(fieldName, fieldValue);
                    }

                    //cs文件名格式
                    string csPath = Path.Combine(curDir.FullName, tblName + this.csCfg.Externion).PathFormat();
                    dataModel.export = csPath;

                    //cs导出包
                    string enumBody = CSBuilder.PackageEnum(tblName.Upper(), classBuilder.ToString());
                    string packageName = this.csCfg.PackageFormat.Format(nameSpace.Upper());
                    string classDesc = CSTemplate.DESC.Format(dataModel.desc);
                    string nameSpaceBody = CSBuilder.PackageNameSpace(packageName, classDesc + enumBody);
                    string usingBody = usingBuilder.ToString();
                    dataModel.txt = usingBody + nameSpaceBody;
                }
            }
        }

        //生成cs 数据表
        private void OutputCSTbl()
        {
            foreach (var nameSpace in this.tblSheetMap.Keys)
            {
                //根据命名空间创建目录
                string exportTo = this.csCfg.ExportTo;
                string dirPath = Path.Combine(exportTo, nameSpace).PathFormat();
                var curDir = Directory.CreateDirectory(dirPath);

                var tblMap = this.tblSheetMap[nameSpace];
                //导出sheet
                foreach (var tblName in tblMap.Keys)
                {
                    var dataModel = tblMap[tblName];
                    var objDataList = dataModel.objDataList;

                    CSBuilder classBuilder = new CSBuilder();
                    CSBuilder usingBuilder = new CSBuilder();
                    for (int i = 0; i < objDataList.Count; ++i)
                    {
                        var objData = objDataList[i];
                        string fieldName = objData.fieldName;
                        string fieldType = objData.fieldType;
                        string fieldDesc = objData.fieldDesc;

                        // 根据数据类型生成cs字段
                        fieldType = ToCSData(fieldType, ref usingBuilder);
                        // 添加字段
                        classBuilder.AddDesc(fieldDesc);
                        classBuilder.AddField(fieldType, fieldName);
                    }

                    //cs文件名格式
                    string csPath = Path.Combine(curDir.FullName, tblName + this.csCfg.Externion).PathFormat();
                    dataModel.export = csPath;

                    //cs导出包
                    string classBody = CSBuilder.PackageClass(tblName.Upper(), classBuilder.ToString());
                    string packageName = this.csCfg.PackageFormat.Format(nameSpace.Upper());
                    string classDesc = CSTemplate.DESC.Format(dataModel.desc);
                    string nameSpaceBody = CSBuilder.PackageNameSpace(packageName, classDesc + classBody);
                    string usingBody = usingBuilder.ToString();
                    dataModel.txt = usingBody + nameSpaceBody;
                }
            }
        }

        private string ToCSData(string fieldType, ref CSBuilder usingBuilder)
        {
            XlsxTypes xlsxTypes = this.xlsxCfg.XlsxTypes;
            if (fieldType == xlsxTypes.Number)
            {
                fieldType = CSTemplate.FLOAT;
            }
            else if (fieldType == xlsxTypes.String)
            {
                fieldType = CSTemplate.STRING;
            }
            else if (fieldType == xlsxTypes.ListNumber)
            {
                fieldType = CSTemplate.LIST.Format(CSTemplate.FLOAT);
            }
            else if (fieldType == xlsxTypes.ListString)
            {
                fieldType = CSTemplate.LIST.Format(CSTemplate.STRING);
            }
            else if (fieldType == xlsxTypes.Xid)
            {
                fieldType = CSTemplate.STRING;
            }
            else if (Regex.IsMatch(fieldType, xlsxTypes.InlineTable))
            {
                var match = Regex.Match(fieldType, this.xlsxCfg.XlsxTypes.InlineTable);
                string inlineTblNameSpace = match.Groups[1].Value;
                string inlineTblName = match.Groups[2].Value;
                string nameSpace = csCfg.PackageFormat.Format(inlineTblNameSpace.Upper());
                usingBuilder.AddUsing(nameSpace);
                fieldType = CSTemplate.PROPERTY.Format(nameSpace, inlineTblName.Upper());
            }
            else
            {
                logger.E("配置表数据类型错误！！！");
            }
            return fieldType;
        }

        private void WriteAllXlsxDataToCS()
        {
            foreach (var (nameSpace, modelMap) in this.objSheetMap)
            {
                foreach (var (tblName, model) in modelMap)
                {
                    WriteTocs(model);
                }
            }

            foreach (var (nameSpace, modelMap) in this.enumSheetMap)
            {
                foreach (var (tblName, model) in modelMap)
                {
                    WriteTocs(model);
                }
            }

            foreach (var (nameSpace, modelMap) in this.tblSheetMap)
            {
                foreach (var (tblName, model) in modelMap)
                {
                    WriteTocs(model);
                }
            }
        }

        private void WriteTocs(XlsxDataModel dataModel)
        {
            string xlsx = dataModel.xlsx;
            string cs = dataModel.export;
            string csTxt = dataModel.txt;
            using (FileStream fs = new FileStream(cs, FileMode.OpenOrCreate, FileAccess.Write))
            {
                fs.SetLength(0);
                byte[] csData = Encoding.UTF8.GetBytes(csTxt);
                fs.Write(csData);
            }
            logger.P("导出完成{0}...".Format(dataModel.export));
        }
    }
}