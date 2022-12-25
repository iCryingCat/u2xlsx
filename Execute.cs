﻿using System.Linq;
using System.Data;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace GFramework.Xlsx
{
    public class Execute
    {
        private static GLogger logger = new GLogger("Execute");

        public static void Main(string[] args)
        {
            XlsxCfg cfg = XlsxExporter.Instance.LoadCfg();

            string sourcePath = cfg.SourcePath;
            string exportPath = cfg.ExportPath;
            logger.P("xlsx路径：{0}".Format(sourcePath));
            if (!Directory.Exists(sourcePath))
                throw new DirectoryNotFoundException(sourcePath);

            logger.P("导出路径：{0}".Format(exportPath));
            if (!Directory.Exists(exportPath))
                throw new DirectoryNotFoundException(exportPath);

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
            logger.P("所有文件导出完成！！！");
            logger.P("按任意键关闭！！！");
            Console.ReadLine();
        }
    }
}