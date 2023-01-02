using System.Globalization;
using System.Text;

namespace GFramework.Xlsx
{
    public class LuaBuilder
    {
        private StringBuilder body = new StringBuilder();

        public static string PackageToTbl(Dictionary<string, XlsxTblItemData> data)
        {
            LuaBuilder itemBuilder = new LuaBuilder();
            foreach (var (key, field) in data)
            {
                itemBuilder.AddObjField(field.fieldName, field.fieldValue);
            }
            return LuaTemplate.TBL.Format(itemBuilder.ToString());
        }

        public static string ToLocalTbl(string tblName, string content)
        {
            return LuaTemplate.LOCAL_TABLE_OBJ.Format(tblName, content);
        }

        public static string Export(string packageName, string packageBody)
        {
            return LuaTemplate.EXPORT.Format(packageBody, packageName);
        }

        public void AddSubBody(string content)
        {
            this.body.AppendLine(content);
        }

        public void AddDesc(string desc)
        {
            this.body.Append(LuaTemplate.DESC.Format(desc));
        }

        public void AddObjField(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.FIELD.Format(key, value));
        }

        public void AddListItem(string key, string value)
        {
            this.body.AppendLine(LuaTemplate.LIST_NUM_ITEM.Format(key, value));
        }

        public override string ToString()
        {
            return this.body.ToString();
        }
    }
}