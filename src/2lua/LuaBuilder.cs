using System.Text;

namespace GFramework.Xlsx
{
    public class LuaBuilder
    {
        private StringBuilder body = new StringBuilder();

        public string ToLocalTbl(string tblName)
        {
            StringBuilder luaTxt = new StringBuilder();
            string localTbl = LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, this.body);
            luaTxt.AppendLine(localTbl);
            luaTxt.AppendLine(LuaTemplate.RETURN.Format(tblName));
            return luaTxt.ToString();
        }

        public void AddSubContent(string content)
        {
            this.body.Append(content);
        }

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

        public override string ToString()
        {
            return this.body.ToString();
        }
    }
}