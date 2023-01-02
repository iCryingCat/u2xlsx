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
            new XlsxExporter().ExecuteExport();
            logger.P("所有文件导出完成！！！");
            logger.P("按任意键关闭！！！");
            Console.ReadLine();
        }
    }
}