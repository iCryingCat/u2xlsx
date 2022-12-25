using System.Text;

namespace GFramework.Xlsx
{
    public class LuaBuilder
    {
        private StringBuilder body = new StringBuilder();

        public override string ToString()
        {
            return this.body.ToString();
        }

        public string ToLocalTbl(string tblName)
        {
            return LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, this.body);
        }

        public void AddSubContent(string content)
        {
            this.body.Append(content);
        }

        public void AddDesc(string desc)
        {
            this.body.AppendLine();
        }

        public void AddObjField(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.FIELD.Format(key, value));
        }

        public void AddListItem(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.LIST_ITEM.Format(key, value));
        }

        public static string ToLuaTable(string[] values)
        {
            StringBuilder sb = new StringBuilder();
            int itemNum = values.Length;
            for (int i = 0; i < itemNum; ++i)
            {
                string value = values[i];
                sb.AppendLine(value);
        }
            return LuaTemplate.TABLE.Format(sb);
        }
}
}