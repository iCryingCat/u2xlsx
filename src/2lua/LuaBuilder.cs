using System.Text;

namespace GFramework.Xlsx
{
    public class LuaBuilder
    {
        private StringBuilder body = new StringBuilder();

        public string ToLocalTbl(string tblName)
        {
            return LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, this.body);
        }

        public string ToTbl()
        {
            return LuaTemplate.TBL.Format(this.body);
        }

        public string ExportPackage(string tblName)
        {
            return LuaTemplate.EXPORT_PACKAGE.Format(this.body, tblName);
        }

        public void AddSubContent(string content)
        {
            this.body.AppendLine(content);
        }

        public void PackLocalTbl(string tblName)
        {
            string localTbl = LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, this.body);
            this.body = new StringBuilder(localTbl);
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