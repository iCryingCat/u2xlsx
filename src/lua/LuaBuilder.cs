using System.Text;

namespace GFramework.Xlsx
{
    public class LuaBuilder
    {
        private StringBuilder body = new StringBuilder();

        public static string ToTbl(Dictionary<string, XlsxTblItemData> data)
        {
            LuaBuilder itemBuilder = new LuaBuilder();
            foreach (var (key, field) in data)
            {
                itemBuilder.AddObjField(field.fieldName, field.fieldValue);
            }
            return LuaTemplate.TBL.Format(itemBuilder.ToString());
        }

        public static string Package(string packageName, string packageBody)
        {
            return LuaTemplate.PACKAGE.Format(packageBody, packageName);
        }

        public static string ToDesc(string desc)
        {
            return LuaTemplate.DESC.Format(desc).Endl();
        }

        public static string ToMultiDesc(string desc)
        {
            return LuaTemplate.MultiDESC.Format(desc).Endl();
        }

        public static string ToTbl(string body)
        {
            return LuaTemplate.TBL.Format(body);
        }

        public void AddSubBody(string content)
        {
            this.body.AppendLine(content);
        }

        public void AddDesc(string desc)
        {
            this.body.AppendLine(LuaTemplate.DESC.Format(desc));
        }

        public void AddObjField(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.FIELD.Format(key, value));
        }

        public void AddListNumItem(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.LIST_NUM_ITEM.Format(key, value));
        }

        public void AddListStrItem(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.LIST_STR_ITEM.Format(key, value));
        }

        public string ToLocalTbl(string tblName)
        {
            return LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, this.body);
        }

        public override string ToString()
        {
            return this.body.ToString();
        }
    }
}