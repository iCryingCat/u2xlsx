using System.Text;

namespace GFramework.Xlsx
{
    public class LuaBuilder
    {
        private StringBuilder body = new StringBuilder();

        public void AddDesc(string desc)
        {
            this.body.AppendLine(LuaTemplate.DESC.Format(desc));
        }

        public void AddObjField(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.FIELD.Format(key, value));
        }

        public void AddListItem(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.LIST_ITEM.Format(key, value));
        }

        public static string ToTbl(string body)
        {
            return LuaTemplate.TBL.Format(body);
        }

        public string ToLocalTbl(string tblName)
        {
            return LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, this.body);
        }

        public static string Package(string packageName, string packageBody)
        {
            return LuaTemplate.EXPORT_PACKAGE.Format(packageBody, packageName);
        }

        public void AddSubBody(string content)
        {
            this.body.AppendLine(content);
        }

        public static string ToDesc(string desc)
        {
            return LuaTemplate.DESC.Format(desc).Endl();
        }

        public static string ToMultiDesc(string desc)
        {
            return LuaTemplate.MultiDESC.Format(desc).Endl();
        }

        public static string ToLocalTbl(string tblName, string tblBody)
        {
            return LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, tblBody);
        }

        public override string ToString()
        {
            return this.body.ToString();
        }
    }
}