using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GFramework.Xlsx
{
    public class CSTemplate
    {
        public const string DESC = "// {0}\n";

        public const string USING = "using {0};\n";

        public const string NAMESPACE = "namespace {0} \n{{\n{1}\n}}\n";

        public const string NUM = "{0}";
        public const string STR = "\"{0}\"";

        public const string NULL = "null";
        public const string FLOAT = "float";
        public const string STRING = "string";
        public const string LIST = "List<{0}>";
        public const string NEW_LIST = "new List<{0}>() {{{1}}}";

        public const string PROPERTY = "{0}.{1}";

        #region 类
        public const string CLASS = "public class {0} \n{{\n{1}\n}}\n";
        public const string ENUM = "public enum {0} \n{{\n{1}\n}}\n";
        public const string EXTENDS = "{0} : {1}";
        #endregion

        #region 字段
        public const string FIELD = "public {0} {1} {{ get; set; }}\n";
        public const string ENUM_ITEM = "{0} = {1},\n";
        #endregion

        #region 方法
        public const string FUNCTION = "public {0} {1} \n{{{2}\n}}\n";
        #endregion
    }
}