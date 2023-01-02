using System.Text;

namespace GFramework.Xlsx
{
    public class CSBuilder
    {
        private StringBuilder body = new StringBuilder();

        public static string PackageNameSpace(string nameSpace, string content)
        {
            return CSTemplate.NAMESPACE.Format(nameSpace, content);
        }

        public static string PackageClass(string className, string content)
        {
            return CSTemplate.CLASS.Format(className, content);
        }

        public static string PackageEnum(string enumName, string content)
        {
            return CSTemplate.ENUM.Format(enumName, content);
        }

        public void AddUsing(string nameSpace)
        {
            this.body.Append(CSTemplate.USING.Format(nameSpace));
        }

        public void AddDesc(string comment)
        {
            this.body.Append(CSTemplate.DESC.Format(comment));
        }

        public void AddSubBody(string sub)
        {
            this.body.Append(sub);
        }

        public void AddField(string fieldType, string fieldName)
        {
            this.body.Append(CSTemplate.FIELD.Format(fieldType, fieldName));
        }

        public void AddEnum(string fieldType, string fieldName)
        {
            this.body.Append(CSTemplate.ENUM_ITEM.Format(fieldType, fieldName));
        }

        public override string ToString()
        {
            return this.body.ToString();
        }
    }
}