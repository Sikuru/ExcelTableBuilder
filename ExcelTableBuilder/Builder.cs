using Microsoft.CSharp;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace Sikuru.ExcelTableBuilder
{
    public partial class Builder
    {
        private string _working_folder;
        private string _namespace_name;

        private TableRawDataStore _table_raw_store;

        private Dictionary<string, Dictionary<string, int>> _enum_kv_dic;
        private Dictionary<string, int> _fk_dic;
        private HashSet<string> _key_check;
        private Assembly _loaded_assembly;

        private readonly string[] _bool_strings = new string[] { "true", "y", "1" };
        private string _class_schema_hash_string;

        public void Build(string working_folder, string bin_data_filename, string namespace_name)
        {
            _working_folder = working_folder;
            _namespace_name = namespace_name;

            if (string.IsNullOrEmpty(bin_data_filename) == true)
            {
                bin_data_filename = $"excel_table_bin.bytes";
            }

            // 엑셀 테이블 Raw 데이터 생성
            var excel_tabale_builder = new ExcelTableBuild();
            _table_raw_store = excel_tabale_builder.Build(working_folder);
            excel_tabale_builder.Dispose();
            if (_table_raw_store == null)
            {
                Trace.WriteLine("엑셀 파일이 없거나 읽을 수 없습니다.");
                return;
            }

            // 기존 코드 로드
            _loaded_assembly = LoadSourceCode();

            // 키 업데이트
            UpdateKeyStore();
            
            // 소스 코드 갱신
            GenerateSourceCode();

            // 새로 생성된 코드 로드
            _loaded_assembly = LoadSourceCode();

            // 갱신된 프리파싱 파일 저장
            var preparsed_file = new FileInfo(Path.Combine(working_folder, "preparsed.bin"));
            using (var bw = new BinaryWriter(preparsed_file.OpenWrite()))
            {
                byte[] preparsed_bytes = TableBinConverter.BinMaker(_table_raw_store);
                byte[] length_bytes = BitConverter.GetBytes(preparsed_bytes.Length);
                bw.Write(length_bytes, 0, length_bytes.Length);
                bw.Write(preparsed_bytes, 0, preparsed_bytes.Length);
            }

            // 테이블 데이터 생성 및 바이너리 저장
            byte[] ts_bin = BuildTableStore();
            var bin_data_file = new FileInfo(bin_data_filename);
            using (var bw = new BinaryWriter(bin_data_file.OpenWrite()))
            {
                byte[] ts_length_bytes = BitConverter.GetBytes((int)ts_bin.Length);
                bw.Write(ts_length_bytes, 0, ts_length_bytes.Length);
                bw.Write(ts_bin, 0, ts_bin.Length);
            }
        }

        private void UpdateKeyStore()
        {
            _fk_dic = new Dictionary<string, int>();
            _key_check = new HashSet<string>();

            if (_loaded_assembly != null)
            {
                // 테이블 유니크 키 enum 복원
                Type table_key_enum = _loaded_assembly.GetType($"{_namespace_name}.TableKeyEnum");
                var key_names = Enum.GetNames(table_key_enum);
                for (int i = 0; i < key_names.Length; ++i)
                {
                    _fk_dic.Add(key_names[i], (int)Enum.Parse(table_key_enum, key_names[i]));
                }
            }

            var key_table_name_check = new HashSet<string>();

            foreach (TableRawData raw in _table_raw_store.RawDataList)
            {
                // 테이블 이름 확인
                if (key_table_name_check.Add(raw.TableName) == false)
                {
                    throw new ApplicationException($"테이블 이름 중복 ; {raw.TableName}");
                }

                int index = raw.FieldTypes.IndexOf(TableFieldType.KEY);
                if (index > -1)
                {
                    foreach (TableRawRecords record_data in raw.Records)
                    {
                        if (record_data.Count() <= index)
                        {
                            throw new ApplicationException($"레코드 인덱스 오류 ; {raw.TableName} {record_data.Count()} <= {index}");
                        }

                        string key = record_data.Get(index);
                        if (string.IsNullOrEmpty(key) || !_key_check.Add(key))
                        {
                            throw new ApplicationException($"키 중복/오류 ; {raw.TableName} {record_data.Get(index)}");
                        }

                        if (_fk_dic.ContainsKey(key) == false)
                        {
                            _fk_dic.Add(key, ++_table_raw_store.LastID);
                        }
                    }
                }
            }

            var removed_keys = new HashSet<string>(_fk_dic.Keys);
            foreach (var key in _key_check)
            {
                removed_keys.Remove(key);
            }

            foreach (var remove_key in removed_keys)
            {
                _fk_dic.Remove(remove_key);
            }
            Trace.WriteLine($"{removed_keys.Count} 개의 키가 삭제됨.");
        }

        private Assembly LoadSourceCode()
        {
            var src_fi = new FileInfo(Path.Combine(_working_folder, "GeneratedTableSource.cs"));
            if (src_fi.Exists == false)
            {
                return null;
            }

            CodeDomProvider code_provider = new CSharpCodeProvider();
            var cp = new CompilerParameters();
            cp.CompilerOptions = "/target:library /optimize";
            cp.GenerateExecutable = false;
            cp.GenerateInMemory = false;
            cp.IncludeDebugInformation = false;

            foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    string location = assembly.Location;
                    if (!string.IsNullOrEmpty(location))
                    {
                        cp.ReferencedAssemblies.Add(location);
                    }
                }
                catch (NotSupportedException)
                {
                }
            }

            var cr = code_provider.CompileAssemblyFromFile(cp, new string[] { src_fi.FullName });
            if (cr.Errors.HasErrors)
            {
                var errors = new StringBuilder("생성된 소스의 컴파일 오류:\r\n");
                foreach (CompilerError error in cr.Errors)
                {
                    errors.AppendFormat("Line {0},{1}\t: {2}\n", error.Line, error.Column, error.ErrorText);
                }
                throw new Exception(errors.ToString());
            }

            var loaded_assembly = AppDomain.CurrentDomain.Load(cr.CompiledAssembly.GetName());
            //foreach (Type type in loaded_assembly.GetTypes())
            //{
            //	Console.WriteLine(type.FullName);
            //}

            return loaded_assembly;
        }
    }
}
