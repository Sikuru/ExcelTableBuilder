using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sikuru.ExcelTableBuilder
{
    public partial class Builder
    {
        private string _working_folder;
        private string _namespace_name;

        private TableRawDataStore _table_raw_store;

        private Dictionary<string, Dictionary<string, int>> _enum_kv_dic;
        private Dictionary<string, HashSet<int>> _fk_int_dic;
        private Dictionary<string, HashSet<string>> _fk_string_dic;

        private readonly string[] _bool_strings = new string[] { "true", "y", "1" };
        private int? _default_enum_id;
        private string _class_schema_hash_string;

        public void Build(string working_folder, string bin_data_filename, string namespace_name)
        {
            _working_folder = working_folder;
            _namespace_name = namespace_name;

            if (string.IsNullOrEmpty(bin_data_filename) == true)
            {
                bin_data_filename = $"excel_table_bin_{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.bytes";
            }

            var bin_fi = new FileInfo(Path.Combine(working_folder, bin_data_filename));

            // 엑셀 테이블 Raw 데이터 생성
            var excel_tabale_builder = new ExcelTableBuild();
            _table_raw_store = excel_tabale_builder.Build(working_folder, bin_data_filename);
            excel_tabale_builder.Dispose();
            if (_table_raw_store == null)
            {
                Trace.WriteLine("엑셀 파일이 없거나 읽을 수 없습니다.");
                return;
            }
        }
    }
}
