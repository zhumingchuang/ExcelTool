using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using UnityEngine;

namespace ExcelTool
{
    public class Vector3Obj
    {
        public float x { get; set; }
        public float y { get; set; }
        public float z { get; set; }
        public Vector3Obj(float x, float y, float z)
        {
            this.x = x;
            this.y = y;
            this.z = z;
        }
    }
    public class Vector2Obj
    {
        public float x { get; set; }
        public float y { get; set; }
        public Vector2Obj(float x, float y)
        {
            this.x = x;
            this.y = y;
        }
    }


    class ExcelUtility
    {


        private static List<string> dataType = new List<string>
        {
            "int",
            "float",
            "long",
            "string",
            "vector2",
            "vector3",
            "vector4"
        };

        private static Dictionary<string, Type> mNameToType = new Dictionary<string, Type>();

        static int Round(float val) { return Mathf.RoundToInt(val); }

        static int Round(double val) { return (int)Math.Round(val); }

        /// <summary>
        /// 名字获取类型
        /// </summary>
        public static Type NameToType(string name)
        {
            Type type = null;
            if (!mNameToType.TryGetValue(name, out type) || type == null)
            {
                if (name == "string") type = typeof(string);
                else if (name == "byte") type = typeof(byte);
                else if (name == "short") type = typeof(Int16);
                else if (name == "int") type = typeof(Int32);
                else if (name == "float") type = typeof(float);
                else if (name == "double") type = typeof(double);
                else if (name == "long") type = typeof(Int64);
                else if (name == "ushort") type = typeof(UInt16);
                else if (name == "uint") type = typeof(UInt32);
                else if (name == "ulong") type = typeof(UInt64);
                else if (name == "vector2") type = typeof(Vector2);
                else if (name == "vector3") type = typeof(Vector3);
                else if (name == "vector4") type = typeof(Vector4);

                mNameToType[name] = type;
            }
            return type;
        }

        /// <summary>
        /// 将指定类型强转需要类型
        /// </summary>
        public static object CastValue(object value, Type desiredType)
        {
            if (value == null) return null;
            var valueType = value.GetType();
            if (valueType == desiredType) return value;
            if (desiredType.IsAssignableFrom(valueType)) return value;

            if (valueType == typeof(int))
            {
                if (desiredType == typeof(byte)) return (byte)(int)value;
                if (desiredType == typeof(short)) return (short)(int)value;
                if (desiredType == typeof(ushort)) return (ushort)(int)value;
                if (desiredType == typeof(uint)) return (uint)(int)value;
                if (desiredType == typeof(float)) return (float)(int)value;
                if (desiredType == typeof(double)) return (double)(int)value;
                if (desiredType == typeof(long)) return (long)(int)value;
                if (desiredType == typeof(ulong)) return (ulong)(int)value;
                if (desiredType == typeof(UInt32)) return (UInt32)(int)value;
                if (desiredType == typeof(Vector3)) return new Vector3((int)value, (int)value, (int)value);
                if (desiredType == typeof(Vector2)) return new Vector2((int)value, (int)value);
                if (desiredType == typeof(LayerMask)) return (LayerMask)(int)value;
                if (desiredType == typeof(string)) return value.ToString();
            }
            else if (valueType == typeof(float))
            {
                if (desiredType == typeof(byte)) return (byte)Round((float)value);
                if (desiredType == typeof(short)) return (short)Round((float)value);
                if (desiredType == typeof(ushort)) return (ushort)Round((float)value);
                if (desiredType == typeof(int)) return Round((float)value);
                if (desiredType == typeof(uint)) return (uint)Round((float)value);
                if (desiredType == typeof(double)) return (double)(float)value;
                if (desiredType == typeof(long)) return (long)Round((float)value);
                if (desiredType == typeof(Vector3)) return new Vector3((float)value, (float)value, (float)value);
                if (desiredType == typeof(Vector2)) return new Vector2((float)value, (float)value);
            }
            else if (valueType == typeof(double))
            {
                if (desiredType == typeof(byte)) return (byte)Round((double)value);
                if (desiredType == typeof(short)) return (short)Round((double)value);
                if (desiredType == typeof(ushort)) return (ushort)Round((double)value);
                if (desiredType == typeof(int)) return Round((double)value);
                if (desiredType == typeof(uint)) return (uint)Round((double)value);
                if (desiredType == typeof(float)) return (float)(double)value;
                if (desiredType == typeof(long)) return (long)Math.Round((double)value);
                if (desiredType == typeof(Vector2)) return new Vector2((float)(double)value, (float)(double)value);
                if (desiredType == typeof(Vector3)) return new Vector3((float)(double)value, (float)(double)value, (float)(double)value);
                if (desiredType == typeof(string)) return value.ToString();
            }
            else if (valueType == typeof(long))
            {
                if (desiredType == typeof(byte)) return (byte)(long)value;
                if (desiredType == typeof(short)) return (short)(long)value;
                if (desiredType == typeof(ushort)) return (ushort)(long)value;
                if (desiredType == typeof(int)) return (int)(long)value;
                if (desiredType == typeof(uint)) return (uint)(long)value;
                if (desiredType == typeof(ulong)) return (ulong)(long)value;
                if (desiredType == typeof(double)) return (double)(long)value;
                if (desiredType == typeof(DateTime)) return new DateTime((long)value, DateTimeKind.Utc);
            }
            else if (valueType == typeof(ulong))
            {
                if (desiredType == typeof(byte)) return (byte)(ulong)value;
                if (desiredType == typeof(short)) return (short)(ulong)value;
                if (desiredType == typeof(ushort)) return (ushort)(ulong)value;
                if (desiredType == typeof(int)) return (int)(ulong)value;
                if (desiredType == typeof(uint)) return (uint)(ulong)value;
                if (desiredType == typeof(long)) return (long)(ulong)value;
                if (desiredType == typeof(double)) return (double)(ulong)value;
                if (desiredType == typeof(DateTime)) return new DateTime((long)(ulong)value, DateTimeKind.Utc);
            }
            else if (valueType == typeof(ushort))
            {
                if (desiredType == typeof(byte)) return (byte)(ushort)value;
                if (desiredType == typeof(short)) return (short)(ushort)value;
                if (desiredType == typeof(int)) return (int)(ushort)value;
                if (desiredType == typeof(uint)) return (uint)(ushort)value;
                if (desiredType == typeof(long)) return (long)(ushort)value;
                if (desiredType == typeof(ulong)) return (ulong)(ushort)value;
                if (desiredType == typeof(double)) return (double)(ushort)value;
            }
            else if (valueType == typeof(byte))
            {
                if (desiredType == typeof(short)) return (short)(byte)value;
                if (desiredType == typeof(ushort)) return (ushort)(byte)value;
                if (desiredType == typeof(int)) return (int)(byte)value;
                if (desiredType == typeof(uint)) return (uint)(byte)value;
                if (desiredType == typeof(long)) return (long)(byte)value;
                if (desiredType == typeof(ulong)) return (ulong)(byte)value;
                if (desiredType == typeof(double)) return (double)(byte)value;
            }
            else if (valueType == typeof(string))
            {
                try
                {
                    if (desiredType == typeof(short)) return Convert.ToInt16(value);
                    if (desiredType == typeof(ushort)) return Convert.ToUInt16(value);
                    if (desiredType == typeof(int)) return Convert.ToInt32(value);
                    if (desiredType == typeof(uint)) return Convert.ToUInt32(value);
                    if (desiredType == typeof(long)) return Convert.ToInt64(value);
                    if (desiredType == typeof(ulong)) return Convert.ToUInt64(value);
                    if (desiredType == typeof(double)) return Convert.ToDouble(value);
                    if (desiredType == typeof(float)) return Convert.ToSingle(value);
                    if (desiredType == typeof(Vector2) || desiredType == typeof(Vector3))
                    {
                        string valueStr = value.ToString();
                        if (valueStr.StartsWith("[") && valueStr.EndsWith("]"))
                        {
                            valueStr = valueStr.Substring(1, valueStr.Length - 2);
                            var vector = valueStr.Split(',');
                            if (desiredType == typeof(Vector3))
                            {
                                float x, y, z;
                                if (float.TryParse(vector[0], out x) && float.TryParse(vector[1], out y) && float.TryParse(vector[2], out z))
                                {
                                    return new Vector3Obj(x, y, z);
                                }
                            }
                            else if (desiredType == typeof(Vector2))
                            {
                                float x, y;
                                if (float.TryParse(vector[0], out x) && float.TryParse(vector[1], out y))
                                {
                                    return new Vector2Obj(x, y);
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
            return null;
        }

        /// <summary>
        /// 读取Excel
        /// </summary>
        public static DataSet ReadExcel(string excelFile)
        {
            try
            {
                FileStream mStream = File.Open(excelFile, FileMode.Open, FileAccess.Read);
                IExcelDataReader mExcelReader = ExcelReaderFactory.CreateOpenXmlReader(mStream);
                return mExcelReader.AsDataSet();
            }
            catch (Exception)
            {
                Log.Write(Level.ERROR, "Excel 文件打开失败");
                throw;
            }
        }

        /// <summary>
        /// 序列化Excel
        /// </summary>
        public static List<string> SerializeExcelToJson(DataSet dataSet)
        {
            if (dataSet.Tables.Count < 1)
                return null;
            List<string> data = new List<string>();
            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                DataTable tempTable = dataSet.Tables[i];
                if (tempTable.Rows.Count < 1)
                    continue;

                int rowCount = tempTable.Rows.Count;
                int colCount = tempTable.Columns.Count;

                List<Dictionary<string, object>> table = new List<Dictionary<string, object>>();

                int rowIndex = 0;
                bool isType = false;
                Dictionary<int, Type[]> dictIndexType = new Dictionary<int, Type[]>();
                Dictionary<int, int> dictAnalysisType = new Dictionary<int, int>();
                for (int j = 0; j < rowCount; j++)
                {
                    //保存每一行的数据
                    Dictionary<string, object> row = new Dictionary<string, object>();
                    for (int k = 0; k < colCount; k++)
                    {
                        string tempfield = tempTable.Rows[j][k].ToString().ToLower();
                        if (!string.IsNullOrEmpty(tempfield))
                        {
                            //找到一个类型
                            if (dataType.Contains(tempfield))
                            {
                                if (!isType)
                                {
                                    isType = true;
                                    rowIndex = j;
                                }
                                dictAnalysisType.Add(k, 0);
                                dictIndexType.Add(k, new[] { NameToType(tempfield) });
                            }

                            //处理字典
                            if (tempfield.StartsWith("dictionary"))
                            {
                                Match matchKey = Regex.Match(tempfield, "dictionary<(.+),");
                                string key = matchKey.Groups[1].Value;

                                Match matchVal = Regex.Match(tempfield, ",(.+)>");
                                string val = matchVal.Groups[1].Value;

                                if (NameToType(key) != null && NameToType(val) != null)
                                {
                                    if (!isType)
                                    {
                                        isType = true;
                                        rowIndex = j;
                                    }
                                    dictAnalysisType.Add(k, 1);
                                    dictIndexType.Add(k, new[] { NameToType(key), NameToType(val) });
                                }
                                else
                                {
                                    Log.Write(Level.ERROR, string.Format("{0}字典类型错误", tempfield));
                                }
                            }

                            //处理列表
                            if (tempfield.StartsWith("list"))
                            {
                                Match matchKey = Regex.Match(tempfield, "list<(.+)>");
                                string type = matchKey.Groups[1].Value;
                                if (NameToType(type) != null)
                                {
                                    if (!isType)
                                    {
                                        isType = true;
                                        rowIndex = j;
                                    }
                                    dictAnalysisType.Add(k, 2);
                                    dictIndexType.Add(k, new[] { NameToType(type) });
                                }
                                else
                                {
                                    Log.Write(Level.ERROR, string.Format("{0}List类型错误", tempfield));
                                }
                            }

                            //处理数组
                            if (tempfield.EndsWith("[]"))
                            {
                                string type = tempfield.Substring(0, tempfield.Length - 2);
                                if (NameToType(type) != null)
                                {
                                    if (!isType)
                                    {
                                        isType = true;
                                        rowIndex = j;
                                    }
                                    dictAnalysisType.Add(k, 2);
                                    dictIndexType.Add(k, new[] { NameToType(type) });
                                }
                                else
                                {
                                    Log.Write(Level.ERROR, string.Format("{0}List类型错误", tempfield));
                                }
                            }

                            if (isType && j > rowIndex + 1 && dictIndexType.ContainsKey(k))
                            {
                                var field = tempTable.Rows[rowIndex + 1][k].ToString();
                                if (dictAnalysisType[k] == 0)
                                {
                                    var value = CastValue(tempTable.Rows[j][k], dictIndexType[k][0]);
                                    if (value == null)
                                    {
                                        Log.Write(Level.ERROR, string.Format("{0}是: {1}类型不等于标记类型{2}", tempTable.Rows[j][k], tempTable.Rows[j][k].GetType(), dictIndexType[k][0]));
                                        return null;
                                    }
                                    //Key-Value对应
                                    row[field] = value;
                                }
                                else if (dictAnalysisType[k] == 1)
                                {
                                    Regex r = new Regex(@"\(.*?\)");
                                    var ms = r.Matches(tempTable.Rows[j][k].ToString());
                                    Dictionary<string, string> tempDict = new Dictionary<string, string>();
                                    foreach (Match item in ms)
                                    {
                                        string value = item.Value.Trim('(', ')');
                                        //解析字典类型
                                        var dictData = value.Split(',');
                                        var key = CastValue(dictData[0], dictIndexType[k][0]);
                                        var val = CastValue(dictData[1], dictIndexType[k][1]);
                                        if (key != null && val != null)
                                        {
                                            if (!tempDict.ContainsKey(key.ToString()))
                                            {
                                                tempDict.Add(key.ToString(), val.ToString());
                                            }
                                            else
                                            {
                                                Log.Write(Level.ERROR, string.Format("重复的Key{0}", key));
                                            }
                                        }
                                        else
                                        {
                                            if (key == null)
                                            {
                                                Log.Write(Level.ERROR, string.Format("数据{0}解析:key{1}不等于标记类型{2}", tempTable.Rows[j][k], dictData[0], dictIndexType[k][0]));
                                            }
                                            if (val == null)
                                            {
                                                Log.Write(Level.ERROR, string.Format("数据{0}解析:key{1}不等于标记类型{2}", tempTable.Rows[j][k], dictData[1], dictIndexType[k][1]));
                                            }
                                        }
                                    }
                                    //Key-Value对应
                                    row[field] = tempDict;
                                }
                                else if (dictAnalysisType[k] == 2)
                                {
                                    string dataStr = tempTable.Rows[j][k].ToString();
                                    if (dataStr.StartsWith("[") && dataStr.EndsWith("]"))
                                    {
                                        Regex r = new Regex(@"\[.*?\]");
                                        var ms = r.Matches(dataStr);
                                        List<object> dataList = new List<object>();
                                        foreach (Match item in ms)
                                        {
                                            var val = CastValue(item.Value, dictIndexType[k][0]);
                                            if (val != null)
                                            {
                                                dataList.Add(val);
                                            }
                                            else
                                            {
                                                Log.Write(Level.ERROR, string.Format("数组成员{0}无法解析为{1}", item.Value, dictIndexType[k][0]));
                                            }
                                        }
                                        //Key-Value对应
                                        row[field] = dataList;
                                    }
                                    else
                                    {
                                        var itemData = dataStr.ToString().Split(',');
                                        List<string> listData = new List<string>();
                                        foreach (var item in itemData)
                                        {
                                            var val = CastValue(item, dictIndexType[k][0]);
                                            listData.Add(val.ToString());
                                            if (val == null)
                                            {
                                                Log.Write(Level.ERROR, string.Format("List{0}中数据{1}不等于标记类型{2}", tempTable.Rows[j][k], item, dictIndexType[k][0]));
                                            }
                                        }
                                        //Key-Value对应
                                        row[field] = listData;
                                    }
                                }
                            }
                        }
                    }
                    //添加到表数据中
                    if (row.Count > 0)
                        table.Add(row);
                }
                string json = JsonConvert.SerializeObject(table, Newtonsoft.Json.Formatting.Indented);
                data.Add(json);
            }
            return data;
        }

        /// <summary>
        /// 序列化Excel
        /// </summary>
        public static List<StringBuilder> SerializeExcelToCSV(DataSet dataSet)
        {
            if (dataSet.Tables.Count < 1)
                return null;
            List<StringBuilder> data = new List<StringBuilder>();
            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                DataTable tempTable = dataSet.Tables[i];
                if (tempTable.Rows.Count < 1)
                    continue;

                int rowCount = tempTable.Rows.Count;
                int colCount = tempTable.Columns.Count;

                int rowIndex = 0;
                bool isType = false;
                Dictionary<int, Type[]> dictIndexType = new Dictionary<int, Type[]>();
                Dictionary<int, int> dictAnalysisType = new Dictionary<int, int>();
                //保存每一行的数据
                StringBuilder stringBuilder = new StringBuilder();
                for (int j = 0; j < rowCount; j++)
                {
                    bool isNull = true;
                    for (int k = 0; k < colCount; k++)
                    {
                        string tempfield = tempTable.Rows[j][k].ToString().ToLower();
                        if (!string.IsNullOrEmpty(tempfield))
                        {
                            //找到一个类型
                            if (dataType.Contains(tempfield))
                            {
                                if (!isType)
                                {
                                    isType = true;
                                    rowIndex = j;
                                }
                                dictAnalysisType.Add(k, 0);
                                dictIndexType.Add(k, new[] { NameToType(tempfield) });
                            }

                            //处理字典
                            if (tempfield.StartsWith("dictionary"))
                            {
                                Match matchKey = Regex.Match(tempfield, "dictionary<(.+),");
                                string key = matchKey.Groups[1].Value;

                                Match matchVal = Regex.Match(tempfield, ",(.+)>");
                                string val = matchVal.Groups[1].Value;

                                if (NameToType(key) != null && NameToType(val) != null)
                                {
                                    if (!isType)
                                    {
                                        isType = true;
                                        rowIndex = j;
                                    }
                                    dictAnalysisType.Add(k, 1);
                                    dictIndexType.Add(k, new[] { NameToType(key), NameToType(val) });
                                }
                                else
                                {
                                    Log.Write(Level.ERROR, string.Format("{0}字典类型错误", tempfield));
                                }
                            }

                            //处理列表
                            if (tempfield.StartsWith("list"))
                            {
                                Match matchKey = Regex.Match(tempfield, "list<(.+)>");
                                string type = matchKey.Groups[1].Value;
                                if (NameToType(type) != null)
                                {
                                    if (!isType)
                                    {
                                        isType = true;
                                        rowIndex = j;
                                    }
                                    dictAnalysisType.Add(k, 2);
                                    dictIndexType.Add(k, new[] { NameToType(type) });
                                }
                                else
                                {
                                    Log.Write(Level.ERROR, string.Format("{0}List类型错误", tempfield));
                                }
                            }

                            //处理数组
                            if (tempfield.EndsWith("[]"))
                            {
                                string type = tempfield.Substring(0, tempfield.Length - 2);
                                if (NameToType(type) != null)
                                {
                                    if (!isType)
                                    {
                                        isType = true;
                                        rowIndex = j;
                                    }
                                    dictAnalysisType.Add(k, 2);
                                    dictIndexType.Add(k, new[] { NameToType(type) });
                                }
                                else
                                {
                                    Log.Write(Level.ERROR, string.Format("{0}List类型错误", tempfield));
                                }
                            }

                            if (j == rowIndex && dictIndexType.ContainsKey(k))
                            {
                                var field = tempTable.Rows[rowIndex + 1][k].ToString();
                                stringBuilder.Append(field + ",");
                                isNull = false;
                            }

                            if (isType && j > rowIndex + 1 && dictIndexType.ContainsKey(k))
                            {
                                if (dictAnalysisType[k] == 0)
                                {
                                    isNull = false;
                                    var value = CastValue(tempTable.Rows[j][k], dictIndexType[k][0]);
                                    if (value == null)
                                    {
                                        Log.Write(Level.ERROR, string.Format("{0}是: {1}类型不等于标记类型{2}", tempTable.Rows[j][k], tempTable.Rows[j][k].GetType(), dictIndexType[k][0]));
                                        return null;
                                    }
                                    stringBuilder.Append(tempTable.Rows[j][k] + ",");
                                }
                                else if (dictAnalysisType[k] == 1)
                                {
                                    Regex r = new Regex(@"\(.*?\)");
                                    var ms = r.Matches(tempTable.Rows[j][k].ToString());
                                    Dictionary<string, string> tempDict = new Dictionary<string, string>();
                                    foreach (Match item in ms)
                                    {
                                        string value = item.Value.Trim('(', ')');
                                        //解析字典类型
                                        var dictData = value.Split(',');
                                        var key = CastValue(dictData[0], dictIndexType[k][0]);
                                        var val = CastValue(dictData[1], dictIndexType[k][1]);
                                        if (key != null && val != null)
                                        {
                                            if (!tempDict.ContainsKey(key.ToString()))
                                            {
                                                tempDict.Add(key.ToString(), val.ToString());
                                            }
                                            else
                                            {
                                                Log.Write(Level.ERROR, string.Format("重复的Key{0}", key));
                                            }
                                        }
                                        else
                                        {
                                            if (key == null)
                                            {
                                                Log.Write(Level.ERROR, string.Format("数据{0}解析:key{1}不等于标记类型{2}", tempTable.Rows[j][k], dictData[0], dictIndexType[k][0]));
                                            }
                                            if (val == null)
                                            {
                                                Log.Write(Level.ERROR, string.Format("数据{0}解析:key{1}不等于标记类型{2}", tempTable.Rows[j][k], dictData[1], dictIndexType[k][1]));
                                            }
                                        }
                                    }
                                    stringBuilder.Append(JsonConvert.SerializeObject(tempDict) + ",");
                                }
                                else if (dictAnalysisType[k] == 2)
                                {
                                    string dataStr = tempTable.Rows[j][k].ToString();
                                    if (dataStr.StartsWith("[") && dataStr.EndsWith("]"))
                                    {
                                        Regex r = new Regex(@"\[.*?\]");
                                        var ms = r.Matches(dataStr);
                                        List<object> dataList = new List<object>();
                                        foreach (Match item in ms)
                                        {
                                            var val = CastValue(item.Value, dictIndexType[k][0]);
                                            if (val != null)
                                            {
                                                dataList.Add(val);
                                            }
                                            else
                                            {
                                                Log.Write(Level.ERROR, string.Format("数组成员{0}无法解析为{1}", item.Value, dictIndexType[k][0]));
                                            }
                                        }
                                        stringBuilder.Append(JsonConvert.SerializeObject(dataList) + ",");
                                    }
                                    else
                                    {
                                        var itemData = dataStr.ToString().Split(',');
                                        List<string> listData = new List<string>();
                                        foreach (var item in itemData)
                                        {
                                            var val = CastValue(item, dictIndexType[k][0]);
                                            if (val == null)
                                            {
                                                Log.Write(Level.ERROR, string.Format("List{0}中数据{1}不等于标记类型{2}", tempTable.Rows[j][k], item, dictIndexType[k][0]));
                                                return null;
                                            }
                                            listData.Add(val.ToString());
                                        }
                                        stringBuilder.Append(JsonConvert.SerializeObject(listData) + ",");
                                    }
                                }
                            }
                        }
                    }
                    if (!isNull)
                        stringBuilder.Append("\r\n");
                }
                data.Add(stringBuilder);
            }
            return data;
        }

        /// <summary>
        /// 转换为Json
        /// </summary>
        public static void ConvertToJson(string excelFile, string jsonPath = null, Encoding encoding = null)
        {
            var dataSet = ReadExcel(excelFile);
            var data = SerializeExcelToJson(dataSet);

            if (encoding == null) encoding = Encoding.UTF8;
            if (data != null)
            {
                for (int i = 0; i < data.Count; i++)
                {
                    string path = jsonPath;
                    if (string.IsNullOrEmpty(jsonPath))
                    {
#if UNITY_STANDALONE_WIN
                        path = GetPath(Path.Combine(Application.dataPath, Path.GetFileNameWithoutExtension(excelFile)), ".json");
#else
                        path = GetPath(Path.Combine(System.Windows.Forms.Application.StartupPath, Path.GetFileNameWithoutExtension(excelFile)), ".json");
#endif
                    }
                    WriterData(path, data[i], encoding);
                }
            }
        }

        /// <summary>
        /// 转换为Xml
        /// </summary>
        public static void ConvertToXml(string excelFile, string xmlPath = null, Encoding encoding = null)
        {

        }

        /// <summary>
        /// 转换为CSV
        /// </summary>
        public static void ConvertToCSV(string excelFile, string csvPath = null, Encoding encoding = null)
        {
            var dataSet = ReadExcel(excelFile);
            var data = SerializeExcelToCSV(dataSet);

            if (encoding == null) encoding = Encoding.UTF8;
            if (data != null)
            {
                for (int i = 0; i < data.Count; i++)
                {
                    string path = csvPath;
                    if (string.IsNullOrEmpty(csvPath))
                    {
#if UNITY_STANDALONE_WIN
                        path = GetPath(Path.Combine(Application.dataPath, Path.GetFileNameWithoutExtension(excelFile)), ".csv");
                        Debug.Log(path);
#else
                        path = GetPath(Path.Combine(System.Windows.Forms.Application.StartupPath, Path.GetFileNameWithoutExtension(excelFile)), ".csv");
                        //Debug.Log(path);
#endif
                    }
                    WriterData(path, data[i].ToString(), encoding);
                }
            }
        }

        /// <summary>
        /// 写入数据
        /// </summary>
        public static void WriterData(string path, string data, Encoding encoding)
        {
            using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter textWriter = new StreamWriter(fileStream, encoding))
                {
                    textWriter.Write(data.ToString());
                }
            }
        }

        /// <summary>
        /// 获取路径
        /// </summary>
        public static string GetPath(string path, string suffix)
        {
            if (File.Exists(path + suffix))
            {
                string pathName = System.IO.Path.GetFileNameWithoutExtension(path);

                int index;
                string tempPath = null;
                if (int.TryParse(pathName.Substring(pathName.Length - 1), out index))
                {
                    tempPath = Path.Combine(Path.GetDirectoryName(path), pathName.Substring(0, pathName.Length - 1) + (index + 1) + suffix);
                    return GetPath(tempPath, suffix);
                }
                else
                {
                    tempPath = Path.Combine(Path.GetDirectoryName(path), pathName + "_1");
                    return GetPath(tempPath, suffix);
                }
            }
            return path + suffix;
        }
    }
}
