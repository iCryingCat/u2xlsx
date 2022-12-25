using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GFramework.Xlsx
{
    public class LuaTemplate
    {
        public const string DESC = "-- {0}";
        public const string NIL = "nil";
        public const string OBJ = "{0}";
        public const string STR = "'{0}'";
        public const string TABLE = "{{ {0} }}";
        public const string LOCAL_TABLE = "local {0} = {}";
        public const string LOCAL_TABLE_OBJ = "local {0} = {{ {1} }}";
        public const string FIELD = "{0} = {1},";
        public const string LIST_ITEM = "[{0}] = {1},";
    }
}