using System.Text;

namespace GFramework.Xlsx
{
    public class CSBuilder
    {
        public string Namespace;
        public string desc = string.Empty;
        private StringBuilder usingBody = new StringBuilder();
        private StringBuilder body = new StringBuilder();
        private StringBuilder subBody = new StringBuilder();

        public CSBuilder(string nameSpace)
        {
            this.Namespace = nameSpace;
        }

        public override string ToString()
        {
            string nameSpaceBody = CSTemplate.Namespace.Format(this.Namespace, this.body);
            StringBuilder sb = new StringBuilder();
            sb.Append(usingBody);
            sb.AppendLine(CSTemplate.Desc.Format(this.desc));
            sb.Append(nameSpaceBody);
            return sb.ToString();
        }

        public void AddUsing(string Ns)
        {
            this.usingBody.AppendLine(CSTemplate.Using.Format(Ns));
        }

        public void AddDesc(string comment)
        {
            this.subBody.AppendLine(CSTemplate.Desc.Format(comment));
        }

        public void AddSubClass(string typeName, string baseTypeName = null)
        {
            string classBody = "";
            if (null == baseTypeName)
            {
                classBody = CSTemplate.PublicClass.Format(typeName, this.subBody);
            }
            else
            {
                classBody = CSTemplate.PublicClassWithExtends.Format(typeName, baseTypeName, this.subBody);
            }
            this.body.Append(classBody);
            this.subBody = new StringBuilder();
        }

        public void AddPublicField(string fieldType, string fieldName, string value = null)
        {
            string fieldBody = "";
            if (null == value)
            {
                fieldBody = CSTemplate.PublicField.Format(fieldType, fieldName);
            }
            else
            {
                fieldBody = CSTemplate.PublicFieldWithValue.Format(fieldType, fieldName, value);
            }
            this.subBody.AppendLine(fieldBody);
        }
    }
}