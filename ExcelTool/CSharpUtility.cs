using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    /// <summary>
    /// 生成数据类
    /// </summary>
    public static class CSharpUtility
    {
        public static string CreateData(string className, Dictionary<string, TypeData> dictField)
        {
            StringBuilder classData = new StringBuilder();

            classData.AppendLine("using System.Collections.Generic;");
            classData.AppendLine("using UnityEngine;");
            classData.AppendLine();
            classData.AppendLine(string.Format("public class {0}", className));
            classData.AppendLine("{");

            foreach (var item in dictField)
            {
                var types = item.Value;
                if (types.types.Length > 1)
                {
                    if (item.Value.typeSign == EAnalysisType.Dict)
                    {
                        var typeKey = ExcelUtility.TypeToName(item.Value.types[0]);
                        var typeVal = ExcelUtility.TypeToName(item.Value.types[1]);
                        string dict = string.Format("Dictionary<{0},{1}>", typeKey, typeVal);
                        classData.AppendLine(string.Format("    public {0} {1};", dict, item.Key));
                    }
                }
                else
                {
                    if (item.Value.typeSign == EAnalysisType.Basic)
                    {
                        var type = ExcelUtility.TypeToName(item.Value.types[0]);
                        classData.AppendLine(string.Format("    public {0} {1};", type, item.Key));
                    }
                    if (item.Value.typeSign == EAnalysisType.Array)
                    {
                        var type = ExcelUtility.TypeToName(item.Value.types[0]);
                        classData.AppendLine(string.Format("    public {0}[] {1};", type, item.Key));
                    }
                    if (item.Value.typeSign == EAnalysisType.List)
                    {
                        var type = ExcelUtility.TypeToName(item.Value.types[0]);
                        classData.AppendLine(string.Format("    public List<{0}> {1};", type, item.Key));
                    }
                }
            }

            classData.AppendLine("}");
            Console.WriteLine(classData.ToString());
            return classData.ToString();
        }

    }
}
