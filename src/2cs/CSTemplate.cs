using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GFramework.Xlsx
{
    public class CSTemplate
    {
        #region 注释
        public const string Desc = "// {0}";
        #endregion

        #region 引用命名
        public const string Using = "using {0};";
        #endregion

        #region 命名空间
        public const string Namespace = "namespace {0} {{\n{1}\n}}";
        #endregion

        #region 类
        public const string PublicClass = "public class {0} {{\n{1}\n}}";
        public const string PublicClassWithExtends = "public class {0} : {1} {{\n{2}\n}}";
        #endregion

        #region 字段
        public const string PublicField = "public {0} {1};";
        public const string PublicFieldWithValue = "public {0} {1} = {2};";
        public const string ProtectedField = "protected {0} {1};";
        public const string ProtectedFieldWithValue = "protected {0} {1} = {2};";
        public const string PrivateField = "private {0} {1};";
        public const string PrivateFieldWithValue = "private {0} {1} = {2};";
        #endregion

        #region 方法
        public const string PublicFunction = "public {0} {1} {{\n{2}\n}}";
        #endregion
    }
}