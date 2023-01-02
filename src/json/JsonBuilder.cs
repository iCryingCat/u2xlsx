using System.Text;

namespace GFramework.Xlsx
{
    public class JsonBuilder
    {
        private string body = string.Empty;

        public static string ToLocalTbl(string tblName, string content)
        {
            return JsonTemplate.FIELD.Format(tblName, JsonTemplate.OBJ.Format(content.ToString()));
        }

        public static string ToTbl(Dictionary<string, XlsxTblItemData> data)
        {
            JsonBuilder itemBuilder = new JsonBuilder();
            foreach (var (key, field) in data)
            {
                itemBuilder.AddObjField(field.fieldName, field.fieldValue);
            }
            return JsonTemplate.OBJ.Format(itemBuilder.ToString());
        }

        public static string Package(string content)
        {
            return JsonTemplate.OBJ.Format(content);
        }

        public void AddObjField(string key, string value)
        {
            string json = JsonTemplate.FIELD.Format(key, value);
            if (string.IsNullOrEmpty(this.body))
                this.body = json.Endl();
            else
                this.body = string.Join(',', this.body, json).Endl();
        }

        public override string ToString()
        {
            return this.body.ToString();
        }
    }
}