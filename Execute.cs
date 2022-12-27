using System.Linq;
using System.Data;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace GFramework.Xlsx
{
    public class Execute
    {
        private static GLogger logger = new GLogger("Execute");

        public static void Main(string[] args)
        {
            ExecuteExport();
            logger.P("所有文件导出完成！！！");
            logger.P("按任意键关闭！！！");
            Console.ReadLine();
        }

        private static void ExecuteExport()
        {
            XlsxCfg cfg = XlsxExporter.Instance.LoadCfg();

            string sourcePath = cfg.SourcePath;
            string exportPath = cfg.ExportPath;
            logger.P("xlsx路径：{0}".Format(sourcePath));
            if (!Directory.Exists(sourcePath))
                throw new DirectoryNotFoundException(sourcePath);

            logger.P("导出路径：{0}".Format(exportPath));
            if (!Directory.Exists(exportPath))
                Directory.CreateDirectory(exportPath);

            string[] flags = cfg.ExportFlags.Split('|');
            foreach (string flag in flags)
            {
                switch (flag)
                {
                    case "lua":
                        logger.P("导出lua...");
                        new LuaExporter().ExportToLua(sourcePath, exportPath);
                        break;
                }
            }
        }
    }
}