using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Sikuru.ExcelTableBuilder
{
    public partial class Builder
    {
        private byte[] BuildTableStore()
        {
            var start_time = DateTime.UtcNow;
            var tdb_ms = new MemoryStream();

            Type type_list = typeof(List<>);
            Type type_table_store = _loaded_assembly.GetType($"{_namespace_name}.TableStore");
            var ts_pi_dic = type_table_store.GetProperties().ToDictionary(x => x.Name);
            object object_ts = Activator.CreateInstance(type_table_store);

            Type type_tablemap = _loaded_assembly.GetType($"{_namespace_name}.TableMap`1");

            foreach (var raw_data in _table_raw_store.RawDataList)
            {
                int key_field_index = raw_data.FieldTypes.IndexOf(TableFieldType.KEY);
                int key_enum_field_index = raw_data.FieldTypes.IndexOf(TableFieldType.KEY_ENUM);
                if (key_field_index == -1 && key_enum_field_index == -1)
                {
                    continue;
                }

                string enum_table_name = null;
                Type type_enum = null;
                Type typed_tablemap = null;

                object object_tablemap = null;
                if (key_field_index > -1)
                {
                    typed_tablemap = type_tablemap.MakeGenericType(typeof(int));
                }
                else if (key_enum_field_index > -1)
                {
                    int index_cm = raw_data.FieldTypeNames[key_enum_field_index].IndexOf(':');
                    enum_table_name = raw_data.FieldTypeNames[key_enum_field_index].Substring(index_cm + 1);
                    type_enum = _loaded_assembly.GetType($"{_namespace_name}.{enum_table_name}");

                    typed_tablemap = type_tablemap.MakeGenericType(type_enum);
                }

                object_tablemap = Activator.CreateInstance(typed_tablemap, null);

                Type type_class = _loaded_assembly.GetType($"{_namespace_name}.{raw_data.TableName}");

                for (int row = 0; row < raw_data.Records.Count; ++row)
                {
                    object row_object = BuildTableClassObject(raw_data, type_class, raw_data.Records[row]);
                    byte[] row_bytes = TableBinConverter.BinMaker(type_class, row_object);

                    object list_position = typed_tablemap.GetProperty("Position").GetValue(object_tablemap);
                    if (list_position == null)
                    {
                        list_position = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(int) }), null);
                    }

                    if (key_field_index > -1)
                    {
                        int key = _fk_dic[(raw_data.Records[row].Get(key_field_index))];

                        object list_key = typed_tablemap.GetProperty("Key").GetValue(object_tablemap);
                        if (list_key == null)
                        {
                            list_key = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(int) }), null);
                        }

                        ((IList)list_key).Add(key);
                        typed_tablemap.GetProperty("Key").SetValue(object_tablemap, list_key);
                    }
                    else if (key_enum_field_index > -1)
                    {
                        object key = Enum.ToObject(type_enum, _enum_kv_dic[enum_table_name][raw_data.Records[row].Get(key_enum_field_index)]);

                        object list_key = typed_tablemap.GetProperty("Key").GetValue(object_tablemap);
                        if (list_key == null)
                        {
                            list_key = Activator.CreateInstance(type_list.MakeGenericType(new[] { type_enum }), null);
                        }

                        ((IList)list_key).Add(key);
                        typed_tablemap.GetProperty("Key").SetValue(object_tablemap, list_key);
                    }

                    ((IList)list_position).Add((int)tdb_ms.Position);
                    typed_tablemap.GetProperty("Position").SetValue(object_tablemap, list_position);
                    tdb_ms.Write(row_bytes, 0, row_bytes.Length);
                }

                ts_pi_dic[raw_data.TableName].SetValue(object_ts, object_tablemap);
            }

            ts_pi_dic["TableSchemaHashBinary"].SetValue(object_ts, _class_schema_hash_string);
            ts_pi_dic["TableDataBin"].SetValue(object_ts, tdb_ms.ToArray());

            byte[] ts_bin = TableBinConverter.BinMaker(type_table_store, object_ts);
            Console.WriteLine("TableStore Binary: {0}bytes, {1}", ts_bin.Length, (DateTime.UtcNow - start_time).ToString());

            return ts_bin;
        }

        private object BuildTableClassObject(TableRawData raw_data, Type type_class, TableRawRecords raw_record)
        {
            Type type_list = typeof(List<>);
            var pi_dic = type_class.GetProperties().ToDictionary(x => x.Name);

            object object_class = Activator.CreateInstance(type_class);

            for (int col = 0; col < raw_data.FieldNames.Count; ++col)
            {
                PropertyInfo pi;
                if (!pi_dic.TryGetValue(raw_data.FieldNames[col], out pi))
                {
                    continue;
                }

                string value = raw_record.Get(col);

                var field_name_counts = raw_data.FieldNameCounts.Find(x => x.FieldName == raw_data.FieldNames[col]);
                switch (raw_data.FieldTypes[col])
                {
                    case TableFieldType.KEY:
                        {
                            if (string.IsNullOrEmpty(value))
                            {
                                throw new InvalidOperationException(string.Format("키 값이 비어있습니다. ; {0}", raw_data.TableName));
                            }

                            pi.SetValue(object_class, _fk_dic[value], null);
                        }
                        break;

                    case TableFieldType.KEY_ENUM:
                        {
                            if (string.IsNullOrEmpty(value))
                            {
                                throw new InvalidOperationException(string.Format("키 값이 비어있습니다. ; {0}", raw_data.TableName));
                            }

                            if (value != null)
                            {
                                var test = new Regex(@"^[a-zA-Z0-9_]*$");
                                if (!test.IsMatch(value))
                                {
                                    throw new InvalidOperationException(string.Format("KEY/ID 에서 사용할 수 없는 문자열 ; {0}", value));
                                }
                            }

                            int index_cm = raw_data.FieldTypeNames[col].IndexOf(':');
                            if (index_cm > -1)
                            {
                                string enum_table_name = raw_data.FieldTypeNames[col].Substring(index_cm + 1);
                                Dictionary<string, int> enum_dic;
                                if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                                {
                                    throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 {0}을 찾을 수 없습니다.", enum_table_name));
                                }

                                if (!enum_dic.ContainsKey(value))
                                {
                                    throw new InvalidOperationException(string.Format("연결된 테이블 {0}에 해당하는 키({1})를 찾을 수 없습니다.", enum_table_name, value));
                                }

                                pi.SetValue(object_class, _enum_kv_dic[enum_table_name][value], null);
                            }
                            else
                            {
                                throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 이름을 ENUM_ID:ENUM테이블명으로 설정해야 합니다. {0}", raw_data.TableName));
                            }
                        }
                        break;

                    case TableFieldType.ID:
                        {
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(int) }), null);
                                }

                                int pv = 0;
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pv = _fk_dic[value];
                                }

                                ((IList)list).Add(pv);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pi.SetValue(object_class, _fk_dic[value], null);
                                }
                            }
                        }
                        break;

                    case TableFieldType.STRING:
                        {
                            string pv = value != null ? value : string.Empty;
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(string) }), null);
                                }

                                ((IList)list).Add(pv);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                pi.SetValue(object_class, pv, null);
                            }
                        }
                        break;

                    case TableFieldType.VECTOR:
                        {
                            if (string.IsNullOrEmpty(value))
                            {
                                continue;
                            }

                            Type type_vector = _loaded_assembly.GetType(string.Format("{0}.TdbVector", _namespace_name));
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { type_vector }), null);
                                }

                                object tdb_vector = Activator.CreateInstance(type_vector, null);
                                if (!string.IsNullOrEmpty(value))
                                {
                                    string[] values = value.Split(',');
                                    type_vector.GetProperty("X").SetValue(tdb_vector, values.Length > 0 ? float.Parse(values[0]) : 0.0f, null);
                                    type_vector.GetProperty("Y").SetValue(tdb_vector, values.Length > 1 ? float.Parse(values[1]) : 0.0f, null);
                                    type_vector.GetProperty("Z").SetValue(tdb_vector, values.Length > 2 ? float.Parse(values[2]) : 0.0f, null);
                                }

                                ((IList)list).Add(tdb_vector);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                object tdb_vector = Activator.CreateInstance(type_vector, null);
                                if (!string.IsNullOrEmpty(value))
                                {
                                    string[] values = value.Split(',');
                                    type_vector.GetProperty("X").SetValue(tdb_vector, values.Length > 0 ? float.Parse(values[0]) : 0.0f, null);
                                    type_vector.GetProperty("Y").SetValue(tdb_vector, values.Length > 1 ? float.Parse(values[1]) : 0.0f, null);
                                    type_vector.GetProperty("Z").SetValue(tdb_vector, values.Length > 2 ? float.Parse(values[2]) : 0.0f, null);
                                }

                                pi.SetValue(object_class, tdb_vector, null);
                            }
                        }
                        break;

                    case TableFieldType.INT:
                        {
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(int) }), null);
                                }

                                int pv = 0;
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pv = int.Parse(value);
                                }

                                ((IList)list).Add(pv);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pi.SetValue(object_class, int.Parse(value), null);
                                }
                            }
                        }
                        break;

                    case TableFieldType.INT_LIST:
                        {
                            Type type_intlist = _loaded_assembly.GetType(string.Format("{0}.IntList", _namespace_name));
                            if (field_name_counts.Count > 1)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { type_intlist }), null);
                                }

                                object int_list = Activator.CreateInstance(type_intlist, null);
                                if (!string.IsNullOrEmpty(value))
                                {
                                    object inner_list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(int) }), null);

                                    string[] values = value.Split(',');
                                    foreach (var split_value in values)
                                    {
                                        ((IList)inner_list).Add(int.Parse(split_value));
                                    }

                                    type_intlist.GetProperty("ListValue").SetValue(int_list, inner_list, null);
                                }

                                ((IList)list).Add(int_list);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                object int_list = Activator.CreateInstance(type_intlist, null);
                                if (!string.IsNullOrEmpty(value))
                                {
                                    object inner_list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(int) }), null);

                                    string[] values = value.Split(',');
                                    foreach (var split_value in values)
                                    {
                                        ((IList)inner_list).Add(int.Parse(split_value));
                                    }

                                    type_intlist.GetProperty("ListValue").SetValue(int_list, inner_list, null);
                                }

                                pi.SetValue(object_class, int_list, null);
                            }
                        }
                        break;

                    case TableFieldType.FLOAT:
                        {
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(float) }), null);
                                }

                                float pv = 0.0f;
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pv = float.Parse(value);
                                }

                                ((IList)list).Add(pv);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pi.SetValue(object_class, float.Parse(value), null);
                                }
                            }
                        }
                        break;

                    case TableFieldType.FLOAT_LIST:
                        {
                            Type type_floatlist = _loaded_assembly.GetType(string.Format("{0}.FloatList", _namespace_name));
                            if (field_name_counts.Count > 1)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { type_floatlist }), null);
                                }

                                object float_list = Activator.CreateInstance(type_floatlist, null);
                                if (!string.IsNullOrEmpty(value))
                                {
                                    object inner_list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(float) }), null);

                                    string[] values = value.Split(',');
                                    foreach (var split_value in values)
                                    {
                                        ((IList)inner_list).Add(float.Parse(split_value));
                                    }

                                    type_floatlist.GetProperty("ListValue").SetValue(float_list, inner_list, null);
                                }

                                ((IList)list).Add(float_list);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                object float_list = Activator.CreateInstance(type_floatlist, null);
                                if (!string.IsNullOrEmpty(value))
                                {
                                    object inner_list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(float) }), null);

                                    string[] values = value.Split(',');
                                    foreach (var split_value in values)
                                    {
                                        ((IList)inner_list).Add(float.Parse(split_value));
                                    }

                                    type_floatlist.GetProperty("ListValue").SetValue(float_list, inner_list, null);
                                }

                                pi.SetValue(object_class, float_list, null);
                            }
                        }
                        break;

                    case TableFieldType.BOOL:
                        {
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(bool) }), null);
                                }

                                bool b = false;
                                if (!string.IsNullOrEmpty(value))
                                {
                                    b = !string.IsNullOrEmpty(value) && _bool_strings.Any(value.ToLower().Trim().Contains);
                                }
                                ((IList)list).Add(b);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    bool b = !string.IsNullOrEmpty(value) && _bool_strings.Any(value.ToLower().Trim().Contains);
                                    pi.SetValue(object_class, b, null);
                                }
                            }
                        }
                        break;

                    case TableFieldType.DATETIME:
                        {
                            if (pi.PropertyType.IsGenericType && pi.PropertyType.GetGenericTypeDefinition() == type_list)
                            {
                                object list = pi.GetValue(object_class);
                                if (list == null)
                                {
                                    list = Activator.CreateInstance(type_list.MakeGenericType(new[] { typeof(double) }), null);
                                }

                                double pv = 0.0d;
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pv = double.Parse(value);
                                }

                                ((IList)list).Add(pv);
                                pi.SetValue(object_class, list, null);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(value))
                                {
                                    pi.SetValue(object_class, double.Parse(value), null);
                                }
                            }
                        }
                        break;

                    case TableFieldType.ENUM_KEY:
                        {
                            var test = new Regex(@"^[a-zA-Z0-9_]*$");
                            if (!test.IsMatch(value))
                            {
                                throw new InvalidOperationException(string.Format("KEY/ID 에서 사용할 수 없는 문자열 ; {0}", value));
                            }
                        }
                        break;

                    case TableFieldType.ENUM_VALUE:
                        {
                            // 적재 안함
                        }
                        break;

                    case TableFieldType.ENUM_ID:
                        {
                            bool value_exist = !string.IsNullOrEmpty(value);
                            if (value_exist == false)
                            {
                                throw new InvalidOperationException(string.Format("ENUM_ID 필드를 비워둘 수 없습니다. {0}", raw_data.TableName));
                            }

                            if (field_name_counts.Count > 1)
                            {
                                int index_cm = raw_data.FieldTypeNames[col].IndexOf(':');
                                if (index_cm > -1)
                                {
                                    string enum_table_name = raw_data.FieldTypeNames[col].Substring(index_cm + 1);
                                    Dictionary<string, int> enum_dic;
                                    if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                                    {
                                        throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 {0}을 찾을 수 없습니다.", enum_table_name));
                                    }

                                    Type type_enum = _loaded_assembly.GetType(string.Format("{0}.{1}", _namespace_name, enum_table_name));
                                    object enum_object;

                                    if (value_exist == false)
                                    {
                                        throw new InvalidOperationException(string.Format("ENUM_ID 필드를 비워둘 수 없습니다. {0}", raw_data.TableName));
                                    }
                                    else
                                    {
                                        if (!enum_dic.ContainsKey(value))
                                        {
                                            throw new InvalidOperationException(string.Format("연결된 테이블 {0}에 해당하는 키({1})를 찾을 수 없습니다.", enum_table_name, value));
                                        }

                                        enum_object = Enum.ToObject(type_enum, _enum_kv_dic[enum_table_name][value]);
                                    }

                                    object list = pi.GetValue(object_class);
                                    if (list == null)
                                    {
                                        list = Activator.CreateInstance(type_list.MakeGenericType(new[] { type_enum }), null);
                                    }

                                    ((IList)list).Add(enum_object);
                                    pi.SetValue(object_class, list, null);
                                }
                                else
                                {
                                    throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 이름을 ENUM_ID:ENUM테이블명으로 설정해야 합니다. {0}", raw_data.TableName));
                                }
                            }
                            else
                            {
                                int index_cm = raw_data.FieldTypeNames[col].IndexOf(':');
                                if (index_cm > -1)
                                {
                                    string enum_table_name = raw_data.FieldTypeNames[col].Substring(index_cm + 1);
                                    Dictionary<string, int> enum_dic;
                                    if (!_enum_kv_dic.TryGetValue(enum_table_name, out enum_dic))
                                    {
                                        throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 {0}을 찾을 수 없습니다.", enum_table_name));
                                    }

                                    Type type_enum = _loaded_assembly.GetType(string.Format("{0}.{1}", _namespace_name, enum_table_name));
                                    object enum_object;
                                    if (value_exist == false)
                                    {
                                        throw new InvalidOperationException(string.Format("ENUM_ID 필드를 비워둘 수 없습니다. {0}", raw_data.TableName));
                                    }
                                    else
                                    {
                                        if (!enum_dic.ContainsKey(value))
                                        {
                                            throw new InvalidOperationException(string.Format("{0}에서 연결된 테이블 {1}에 해당하는 키({2})를 찾을 수 없습니다.", raw_data.TableName, enum_table_name, value));
                                        }

                                        enum_object = Enum.ToObject(type_enum, _enum_kv_dic[enum_table_name][value]);
                                    }

                                    pi.SetValue(object_class, enum_object, null);
                                }
                                else
                                {
                                    throw new InvalidOperationException(string.Format("연결된 ENUM 테이블 이름을 ENUM_ID:ENUM테이블명으로 설정해야 합니다. {0}", raw_data.TableName));
                                }
                            }
                        }
                        break;

                    default:
                        {
                            throw new InvalidOperationException("Unknown Data Type");
                        }
                }
            }

            return object_class;
        }
    }
}
