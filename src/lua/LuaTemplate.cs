using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GFramework.Xlsx
{
    public class LuaTemplate
    {
        public const string DESC = "-- {0}\n";
        public const string MULTI_DESC = "--[[\n{0}\n--]]\n";
        public const string NIL = "nil";
        public const string OBJ = "{0}";
        public const string TBL = "{{\n{0}}}";
        public const string STR = "'{0}'";

        public const string LOCAL_TABLE_OBJ = "local {0} = {{\n{1}}}";
        public const string EXPORT = "{0}\nreturn {1}";
        public const string FIELD = "{0} = {1},";
        public const string LIST_NUM_ITEM = "[{0}] = {1},";
    }
}