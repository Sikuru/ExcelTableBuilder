using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Sikuru.ExcelTableBuilder
{
    public partial class Builder
    {
        private void GenerateSourceCode()
        {
            string cs_file = Path.Combine(_working_folder, "GeneratedTableSource.cs");
            Trace.WriteLine($"Generate {cs_file}");

            using (var fs = new StreamWriter(cs_file, false))
            {
                fs.WriteLine("using System;");
                fs.WriteLine("using System.Collections;");
                fs.WriteLine("using System.Collections.Generic;");
                fs.WriteLine("using System.IO;");
                fs.WriteLine("using System.Linq;");
                fs.WriteLine("using System.Reflection;");
                fs.WriteLine("using System.Text;");
                fs.WriteLine();
                fs.WriteLine($"namespace {_namespace_name}");
                fs.WriteLine("{");

                //for (int i = 0; i < _)
                fs.WriteLine("\tpublic enum TableKeyEnum : int");
                fs.WriteLine("\t{");
                foreach (var row in _fk_dic)
                {
                    fs.WriteLine($"\t\t{row.Key} = {row.Value},");
                }
                fs.WriteLine("\t}");
                fs.WriteLine();

                // 테이블당 enum
                _enum_kv_dic = new Dictionary<string, Dictionary<string, int>>();
                for (int i = 0; i < _table_raw_store.RawDataList.Count; ++i)
                {
                    var table = _table_raw_store.RawDataList[i];
                    int enum_field_index = table.FieldTypes.IndexOf(TableFieldType.ENUM_KEY);
                    int enum_value_field_index = table.FieldTypes.IndexOf(TableFieldType.ENUM_VALUE);
                    if (enum_field_index >= 0 && enum_value_field_index >= 0)
                    {
                        fs.WriteLine($"\tpublic enum {table.TableName} : int");
                        fs.WriteLine("\t{");

                        for (int j = 0; j < table.Records.Count; ++j)
                        {
                            var kv = new KeyValuePair<string, int>(table.Records[j].Get(enum_field_index), int.Parse(table.Records[j].Get(enum_value_field_index)));

                            fs.WriteLine($"\t\t{kv.Key} = {kv.Value},");

                            if (!_enum_kv_dic.ContainsKey(table.TableName))
                            {
                                _enum_kv_dic.Add(table.TableName, new Dictionary<string, int>());
                            }
                            _enum_kv_dic[table.TableName].Add(kv.Key, kv.Value);

                        }
                        fs.WriteLine("\t}");
                        fs.WriteLine();
                    }
                }

                // 어트리뷰트
                fs.WriteLine("#region TdbExtentions");
                fs.WriteLine("\t[AttributeUsage(AttributeTargets.Property)]");
                fs.WriteLine("\tpublic class TableKey : Attribute");
                fs.WriteLine("\t{");
                fs.WriteLine("\t\tprivate bool _is_key = false;");
                fs.WriteLine();
                fs.WriteLine("\t\tpublic TableKey(bool is_key = true)");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\t_is_key = is_key;");
                fs.WriteLine("\t\t}");
                fs.WriteLine("");
                fs.WriteLine("\t\tpublic bool IsTableKey()");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\treturn _is_key;");
                fs.WriteLine("\t\t}");
                fs.WriteLine("\t}");
                fs.WriteLine();

                // 테이블 클래스 베이스
                fs.WriteLine("\tpublic abstract class TableBase : TableBinConvertable");
                fs.WriteLine("\t{");
                fs.WriteLine("\t}");
                fs.WriteLine();

                // INT_LIST, FLOAT_LIST, VECTOR
                fs.WriteLine("\tpublic class IntList : TableBinConvertable");
                fs.WriteLine("\t{");
                fs.WriteLine("\t\tpublic List<int> ListValue { get; set; }");
                fs.WriteLine();
                fs.WriteLine("\t\tpublic List<int> GetListValue()");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\treturn this.ListValue != null ? this.ListValue : new List<int>();");
                fs.WriteLine("\t\t}");
                fs.WriteLine("\t}");
                fs.WriteLine();
                fs.WriteLine("\tpublic class FloatList : TableBinConvertable");
                fs.WriteLine("\t{");
                fs.WriteLine("\t\tpublic List<float> ListValue { get; set; }");
                fs.WriteLine();
                fs.WriteLine("\t\tpublic List<float> GetListValue()");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\treturn this.ListValue != null ? this.ListValue : new List<float>();");
                fs.WriteLine("\t\t}");
                fs.WriteLine("\t}");
                fs.WriteLine();
                fs.WriteLine("\tpublic class TdbVector : TableBinConvertable");
                fs.WriteLine("\t{");
                fs.WriteLine("\t\tpublic float X { get; set; }");
                fs.WriteLine("\t\tpublic float Y { get; set; }");
                fs.WriteLine("\t\tpublic float Z { get; set; }");
                fs.WriteLine("\t}");
                fs.WriteLine();
                fs.WriteLine("\tpublic class TableMap<T> : TableBinConvertable");
                fs.WriteLine("\t{");
                fs.WriteLine("\t\tpublic List<T> Key { get; set; }");
                fs.WriteLine("\t\tpublic List<int> Position { get; set; }");
                fs.WriteLine("\t}");
                fs.WriteLine("#endregion");
                fs.WriteLine();

                // 클래스 정의
                var class_schema_stream = new MemoryStream();
                for (int i = 0; i < _table_raw_store.RawDataList.Count; ++i)
                {
                    var table = _table_raw_store.RawDataList[i];
                    int key_field_index = table.FieldTypes.IndexOf(TableFieldType.KEY);
                    int key_enum_field_index = table.FieldTypes.IndexOf(TableFieldType.KEY_ENUM);
                    if (key_field_index == -1 && key_enum_field_index == -1)
                    {
                        // 키가 없는 테이블 ENUM_KEY 테이블은 스킵
                        continue;
                    }

                    fs.WriteLine($"\tpublic class {table.TableName} : TableBase");
                    fs.WriteLine("\t{");

                    var property_name_check = new HashSet<string>();
                    for (int j = 0; j < table.FieldTypes.Count; ++j)
                    {
                        if (property_name_check.Contains(table.FieldNames[j]))
                        {
                            continue;
                        }
                        property_name_check.Add(table.FieldNames[j]);

                        byte[] class_field_bytes = Encoding.UTF8.GetBytes(table.FieldNames[j]);
                        class_schema_stream.Write(class_field_bytes, 0, class_field_bytes.Length);

                        var field_name_counts = table.FieldNameCounts.Find(x => x.FieldName == table.FieldNames[j]);
                        switch (table.FieldTypes[j])
                        {
                            case TableFieldType.KEY:
                                {
                                    fs.WriteLine("\t\t[TableKey]");
                                    fs.WriteLine($"\t\tpublic int {table.FieldNames[j]}" + @" { get; set; }");
                                }
                                break;
                            case TableFieldType.KEY_ENUM:
                                {
                                    int index_cm = table.FieldTypeNames[j].IndexOf(':');
                                    string enum_table_name = table.FieldTypeNames[j].Substring(index_cm + 1);
                                    Dictionary<string, int> enum_dic;
                                    if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                                    {
                                        throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 {0}을 찾을 수 없습니다. {1}", enum_table_name, table.TableName));
                                    }
                                    fs.WriteLine("\t\t[TableKey]");
                                    fs.WriteLine(string.Format("\t\tpublic {0} {1}", enum_table_name, table.FieldNames[j]) + @" { get; set; }");
                                }
                                break;
                            case TableFieldType.ID:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<int> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic int {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.INT:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<int> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic int {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.INT_LIST:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<IntList> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic IntList {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.FLOAT:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<float> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic float {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.FLOAT_LIST:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<FloatList> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic FloatList {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.BOOL:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<bool> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic bool {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.DATETIME:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<double> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic double {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.STRING:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<string> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic string {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.VECTOR:
                                {
                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<TdbVector> {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic TdbVector {0}", table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                            case TableFieldType.ENUM_ID:
                                {
                                    int index_cm = table.FieldTypeNames[j].IndexOf(':');
                                    string enum_table_name = table.FieldTypeNames[j].Substring(index_cm + 1);
                                    Dictionary<string, int> enum_dic;
                                    if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                                    {
                                        throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 {0}을 찾을 수 없습니다. {1}", enum_table_name, table.TableName));
                                    }

                                    if (field_name_counts.Count > 1)
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic List<{0}> {1}", enum_table_name, table.FieldNames[j]) + @" { get; set; }");
                                    }
                                    else
                                    {
                                        fs.WriteLine(string.Format("\t\tpublic {0} {1}", enum_table_name, table.FieldNames[j]) + @" { get; set; }");
                                    }
                                }
                                break;
                        }
                    }

                    fs.WriteLine("\t}");
                    fs.WriteLine();
                }

                byte[] class_schema_hash = SHA256Managed.Create().ComputeHash(class_schema_stream);
                class_schema_stream.Close();

                // 클래스 스키마 해시
                _class_schema_hash_string = BitConverter.ToString(class_schema_hash).Replace("-", string.Empty);

                // 클래스 스토어2
                fs.WriteLine("#region TableStore3");
                fs.WriteLine("\tpublic class TableStore : TableBinConvertable");
                fs.WriteLine("\t{");
                fs.WriteLine($"\t\tpublic const string TableSchemaHashSource = @\"{_class_schema_hash_string}\";");
                fs.WriteLine("\t\tpublic string TableSchemaHashBinary { get; set; }");
                fs.WriteLine("\t\tpublic byte[] TableDataBin { get; set; }");
                fs.WriteLine();
                for (int i = 0; i < _table_raw_store.RawDataList.Count; ++i)
                {
                    int key_field_index = _table_raw_store.RawDataList[i].FieldTypes.IndexOf(TableFieldType.KEY);
                    int key_enum_field_index = _table_raw_store.RawDataList[i].FieldTypes.IndexOf(TableFieldType.KEY_ENUM);
                    if (key_field_index == -1 && key_enum_field_index == -1)
                    {
                        continue;
                    }

                    if (key_field_index > -1)
                    {
                        fs.WriteLine(string.Format("\t\tpublic TableMap<int> {0} {1}",
                            _table_raw_store.RawDataList[i].TableName, @"{ get; set; }"));
                    }
                    else if (key_enum_field_index > -1)
                    {
                        int index_cm = _table_raw_store.RawDataList[i].FieldTypeNames[key_enum_field_index].IndexOf(':');
                        string enum_table_name = _table_raw_store.RawDataList[i].FieldTypeNames[key_enum_field_index].Substring(index_cm + 1);

                        fs.WriteLine(string.Format("\t\tpublic TableMap<{0}> {1} {2}",
                            enum_table_name, _table_raw_store.RawDataList[i].TableName, @"{ get; set; }"));
                    }
                }
                fs.WriteLine("\t}");
                fs.WriteLine("#endregion");
                fs.WriteLine();

                // 테이블 매니저
                fs.WriteLine("#region TableDataManager");
                fs.WriteLine("\tpublic class TableDataManager");
                fs.WriteLine("\t{");
                fs.WriteLine("\t\tpublic string TableSchemaHashBinary { get; set; }");
                fs.WriteLine();
                for (int i = 0; i < _table_raw_store.RawDataList.Count; ++i)
                {
                    int key_field_index = _table_raw_store.RawDataList[i].FieldTypes.IndexOf(TableFieldType.KEY);
                    int key_enum_field_index = _table_raw_store.RawDataList[i].FieldTypes.IndexOf(TableFieldType.KEY_ENUM);
                    if (key_field_index == -1 && key_enum_field_index == -1)
                    {
                        continue;
                    }

                    if (key_field_index > -1)
                    {
                        fs.WriteLine(string.Format("\t\tpublic TableProxy<int, {0}> {1} {2}",
                            _table_raw_store.RawDataList[i].TableName,
                            _table_raw_store.RawDataList[i].TableName,
                            @"{ get; private set; }"));
                    }

                    if (key_enum_field_index > -1)
                    {
                        int index_cm = _table_raw_store.RawDataList[i].FieldTypeNames[key_enum_field_index].IndexOf(':');
                        string enum_table_name = _table_raw_store.RawDataList[i].FieldTypeNames[key_enum_field_index].Substring(index_cm + 1);
                        Dictionary<string, int> enum_dic;
                        if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                        {
                            throw new InvalidOperationException($"연결된 ENUM 테이블 {enum_table_name}을 찾을 수 없습니다. {_table_raw_store.RawDataList[i].TableName}");
                        }

                        fs.WriteLine(string.Format("\t\tpublic TableProxy<{0}, {1}> {2} {3}",
                            enum_table_name,
                            _table_raw_store.RawDataList[i].TableName,
                            _table_raw_store.RawDataList[i].TableName,
                            @"{ get; private set; }"));
                    }
                }
                fs.WriteLine();
                fs.WriteLine("\t\tprivate TableDataManager(TableStore ts, bool initial_full_parse)");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\tTableSchemaHashBinary = ts.TableSchemaHashBinary;");
                for (int i = 0; i < _table_raw_store.RawDataList.Count; ++i)
                {
                    int key_field_index = _table_raw_store.RawDataList[i].FieldTypes.IndexOf(TableFieldType.KEY);
                    int key_enum_field_index = _table_raw_store.RawDataList[i].FieldTypes.IndexOf(TableFieldType.KEY_ENUM);
                    if (key_field_index == -1 && key_enum_field_index == -1)
                    {
                        continue;
                    }

                    if (key_field_index > -1)
                    {
                        fs.WriteLine(string.Format("\t\t\t{0} = new TableProxy<int, {1}>(ts.{2}, ts.TableDataBin, initial_full_parse);",
                            _table_raw_store.RawDataList[i].TableName,
                            _table_raw_store.RawDataList[i].TableName,
                            _table_raw_store.RawDataList[i].TableName));
                    }

                    if (key_enum_field_index > -1)
                    {
                        int index_cm = _table_raw_store.RawDataList[i].FieldTypeNames[key_enum_field_index].IndexOf(':');
                        string enum_table_name = _table_raw_store.RawDataList[i].FieldTypeNames[key_enum_field_index].Substring(index_cm + 1);
                        Dictionary<string, int> enum_dic;
                        if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                        {
                            throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 {0}을 찾을 수 없습니다. {1}", enum_table_name, _table_raw_store.RawDataList[i].TableName));
                        }

                        fs.WriteLine(string.Format("\t\t\t{0} = new TableProxy<{1}, {2}>(ts.{3}, ts.TableDataBin, initial_full_parse);",
                            _table_raw_store.RawDataList[i].TableName,
                            enum_table_name,
                            _table_raw_store.RawDataList[i].TableName,
                            _table_raw_store.RawDataList[i].TableName));
                    }
                }
                fs.WriteLine("\t\t}");
                fs.WriteLine();
                fs.WriteLine("\t\tpublic static TableDataManager LoadBinaryFile(string ts_bin_file, bool initial_full_parse = false)");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\tbyte[] ts_bin;");
                fs.WriteLine("\t\t\tusing (var bw = new BinaryReader(new FileInfo(ts_bin_file).Open(FileMode.Open, FileAccess.Read)))");
                fs.WriteLine("\t\t\t{");
                fs.WriteLine("\t\t\t\tint length = bw.ReadInt32();");
                fs.WriteLine("\t\t\t\tts_bin = bw.ReadBytes(length);");
                fs.WriteLine("\t\t\t}");
                fs.WriteLine();
                fs.WriteLine("\t\t\tvar ts = TableBinConverter.ClassMaker<TableStore>(ts_bin);");
                fs.WriteLine("\t\t\treturn new TableDataManager(ts, initial_full_parse);");
                fs.WriteLine("\t\t}");
                fs.WriteLine();
                fs.WriteLine("\t\tpublic static TableDataManager LoadBinaryBytes(byte[] ts_bin_bytes, bool initial_full_parse = false)");
                fs.WriteLine("\t\t{");
                fs.WriteLine("\t\t\tbyte[] ts_bin;");
                fs.WriteLine("\t\t\tusing (var bw = new BinaryReader(new MemoryStream(ts_bin_bytes)))");
                fs.WriteLine("\t\t\t{");
                fs.WriteLine("\t\t\t\tint length = bw.ReadInt32();");
                fs.WriteLine("\t\t\t\tts_bin = bw.ReadBytes(length);");
                fs.WriteLine("\t\t\t}");
                fs.WriteLine();
                fs.WriteLine("\t\t\tvar ts = TableBinConverter.ClassMaker<TableStore>(ts_bin);");
                fs.WriteLine("\t\t\treturn new TableDataManager(ts, initial_full_parse);");
                fs.WriteLine("\t\t}");
                fs.WriteLine("\t}");
                fs.WriteLine("#endregion");

                fs.WriteLine();
                fs.WriteLine("#region TableBinConvert");
                fs.Write("\t");
                fs.WriteLine(Properties.Resources.TableBinConverterSource);
                fs.WriteLine("#endregion");

                fs.WriteLine();
                fs.WriteLine("#region TableProxy");
                fs.Write("\t");
                fs.WriteLine(Properties.Resources.TableProxy);
                fs.WriteLine("#endregion");

                // close namespace
                fs.WriteLine("}");
            }
        }
    }
}
