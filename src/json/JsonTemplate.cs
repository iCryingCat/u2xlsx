using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GFramework.Xlsx
{
    public class JsonTemplate
    {
        public const string NIL = "";
        public const string NUM = "{0}";
        public const string STR = "\"{0}\"";
        public const string OBJ = "{{\n{0}}}";
        public const string LIST = "[{0}]";
        public const string FIELD = "\"{0}\" : {1}";
    }
}