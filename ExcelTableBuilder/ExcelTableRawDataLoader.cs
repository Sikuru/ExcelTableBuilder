using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sikuru.ExcelTableBuilder
{
    // 1. 테이블 시그너쳐를 발견할때까지 읽기... 뭐, 이건 전 범위 검색
    // 2. !TABLE 및 테이블 이름
    // 3. !FIELD_NAME 및 !EOF ; end of field
    // 4. !FIELD_TYPE (!KEY, !ID, !ENUM, !STRING, !INT, !FLOAT, !DATETIME) ; !ID는 !KEY 로 다른데서 정의된 ID 값을 사용하는 경우
    // 5. !EOT ; end of table

    public enum TableFieldType
    {
        INVALID,        // 미사용
        KEY,            // 키 타입 선언 (테이블의 키, Enum으로 선언될 수 있는 이름이어야 합니다, 알파벳으로 시작)
        KEY_ENUM,       // Enum 테이블에서 연결된 키
        ID,             // 키를 컬럼의 값으로 사용하는 경우 (다른 테이블 키 참조)
        STRING,         // 문자열 (유니코드)
        INT,            // int
        INT_LIST,       // 셀당 int 리스트 (구분자는 ,를 사용)
        FLOAT,          // float
        FLOAT_LIST,     // 셀당 float 리스트 (구분자는 ,를 사용)
        BOOL,           // 부울 형식. true/false, y/n, 1/0 을 지원합니다. 대소문자를 가리지 않습니다. true/y/1 이외엔 모두 false 입니다.
        DATETIME,       // 날짜 형식. 엑셀 OAD 규격의 날짜 데이터를 double로 가져옵니다.
        VECTOR,         // Vector3 형식
        ENUM_KEY,       // Enum 형태의 키 타입 선언
        ENUM_ID,        // Enum 형태의 키를 컬럼 값으로 사용하는 경우
        ENUM_VALUE,     // Enum 형태의 키의 값
    }

    public class TableRawRecords : TableBinConvertable
    {
        public List<string> RecordData { get; set; }

        public TableRawRecords()
        {
        }

        public TableRawRecords(string[] array)
        {
            RecordData = new List<string>(array);
        }

        public int Count()
        {
            return RecordData.Count;
        }

        public string Get(int index)
        {
            return RecordData[index];
        }
    }

    public class TableRawFieldNameCount : TableBinConvertable
    {
        public string FieldName { get; set; }
        public int Count { get; set; }
    }

    public class TableRawData : TableBinConvertable
    {
        public string TableName { get; set; }
        public string FileName { get; set; }
        public string Hash { get; set; }
        public List<string> FieldNames { get; set; }
        public List<TableRawFieldNameCount> FieldNameCounts { get; set; }
        public List<TableFieldType> FieldTypes { get; set; }
        public List<string> FieldTypeNames { get; set; }
        public List<TableRawRecords> Records { get; set; }
    }

    public class TableRawDataStore : TableBinConvertable
    {
        public List<TableRawData> RawDataList { get; set; }
    }

    struct TableRange
    {
        public int start_row;
        public int start_column;
        public int end_row;
        public int end_column;
    }

    public class ExcelTableBuild : IDisposable, IEnumerable<TableRawData>
    {
        private bool _disposed = false;
        private Excel.Application _excel_app;
        private TableRawDataStore _table_raw_data_store = new TableRawDataStore()
        {
            RawDataList = new List<TableRawData>()
        };

        public const string LABEL_TABLE = "!TABLE";
        public const string LABEL_FIELD_NAME = "!FIELD_NAME";
        public const string LABEL_EOF = "!EOF";
        public const string LABEL_FIELD_TYPE = "!FIELD_TYPE";
        public const string LABEL_EOT = "!EOT";

        public ExcelTableBuild()
        {
        }

        public IEnumerator<TableRawData> GetEnumerator()
        {
            foreach (TableRawData raw in _table_raw_data_store.RawDataList)
            {
                yield return raw;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public TableRawDataStore GetStore()
        {
            return _table_raw_data_store;
        }

        public TableRawDataStore Build(string build_path, bool initial_build)
        {
            //var preparsed = new FileInfo(string.Format(@"{0}\TdbRawData.preparsed", build_path));
            //if (preparsed.Exists && preparsed.Length > 0)
            //{
            //	Trace.WriteLine("Load preparsed data {0}", preparsed.Length);
            //	using (var br = new BinaryReader(preparsed.OpenRead()))
            //	{
            //		int length = br.ReadInt32();
            //		byte[] compressed = br.ReadBytes(length);
            //		byte[] preparsed_bytes = LibLZF.CLZF2.Decompress(compressed);
            //		_table_raw_data_store = TableBinConverter.ClassMaker<TableRawDataStore>(preparsed_bytes);
            //	}
            //}

            var start_time = DateTime.UtcNow;

            _excel_app = new Excel.Application();
            _excel_app.Visible = false;
            _excel_app.DisplayAlerts = false;
            _excel_app.ScreenUpdating = false;

            try
            {
                var sha256 = new SHA256Managed();

                string[] filepaths = Directory.GetFiles(build_path, "*.xlsx", SearchOption.AllDirectories);
                if (filepaths == null || filepaths.Length == 0)
                {
                    return null;
                }

                foreach (string filepath in filepaths)
                {
                    if (File.GetAttributes(filepath).HasFlag(FileAttributes.Hidden))
                    {
                        continue;
                    }

                    string filehash = "";
                    using (var stream = File.OpenRead(filepath))
                    {
                        byte[] hashbytes = sha256.ComputeHash(stream);
                        filehash = BitConverter.ToString(hashbytes).Replace("-", string.Empty);
                    }

                    string filename = Path.GetFileName(filepath);
                    TableRawData raw_data = _table_raw_data_store.RawDataList.Find(x => x.FileName == filename);
                    if (raw_data != null && !string.IsNullOrEmpty(raw_data.Hash))
                    {
                        if (raw_data.Hash.Equals(filehash, StringComparison.OrdinalIgnoreCase))
                        {
                            Trace.WriteLine($"Skipped (Hash Match)... {filename} ({filehash})");
                            continue;
                        }
                        else
                        {
                            var remove_list = _table_raw_data_store.RawDataList.Where(x => x.FileName == filename).ToList();
                            for (int i = 0; i < remove_list.Count; ++i)
                            {
                                _table_raw_data_store.RawDataList.Remove(remove_list[i]);
                                Trace.WriteLine("Removed preparsed... {0}", remove_list[i].TableName);
                            }
                        }
                    }

                    Trace.WriteLine($"\r\nParsing... {filename} ({filehash})");
                    Parse2(filepath, filename, filehash);
                }

                //int preparsed_length = 0;
                //using (var bw = new BinaryWriter(preparsed.OpenWrite()))
                //{
                //	byte[] preparsed_bytes = TableBinConverter.BinMaker(_table_raw_data_store);
                //	byte[] compressed = LibLZF.CLZF2.Compress(preparsed_bytes);

                //	preparsed_length = compressed.Length;

                //	byte[] length_bytes = BitConverter.GetBytes(preparsed_length);
                //	bw.Write(length_bytes, 0, length_bytes.Length);
                //	bw.Write(compressed, 0, compressed.Length);
                //}

                Trace.WriteLine($"\r\nParse Completed. {(DateTime.UtcNow - start_time).ToString()}");
                return _table_raw_data_store;
            }
            finally
            {
                _excel_app.ScreenUpdating = true;
                _excel_app.Quit();
            }
        }

        private void Parse(string filepath, string filename, string hash)
        {
            // open
            var wb = _excel_app.Workbooks.Open(filepath);
            var worksheets = wb.Worksheets;

            Trace.WriteLine($"Worksheets.Count={worksheets.Count}");

            // load all cells
            foreach (Excel.Worksheet current_sheet in worksheets)
            {
                if (current_sheet.Name.IndexOf('#') != -1)
                {
                    continue;
                }

                var range = current_sheet.UsedRange;
                Trace.WriteLine($"ActiveSheet.UsedRange Rows={range.Rows.Count}, Columns={range.Columns.Count} <{current_sheet.Name}>");

                int row_count = range.Rows.Count;
                int column_count = range.Columns.Count;

                if (row_count > 100000 || column_count > 100000)
                {
                    throw new ApplicationException($"COUNT OVERFLOW ; ActiveSheet.UsedRange Rows={range.Rows.Count}, Columns={range.Columns.Count} <{current_sheet.Name}>");
                }

                for (int i = 1; i <= row_count; ++i)
                {
                    for (int j = 1; j <= column_count; ++j)
                    {
                        var cell = current_sheet.Cells[i, j];
                        if (cell.Value2 != null)
                        {
                            string cell_string = cell.Value2.ToString();

                            if (cell_string.Equals(LABEL_TABLE))
                            {
                                // 테이블 디텍트
                                TableRange table_range = DetectTable(current_sheet, i, j);

                                // raw data 생성
                                TableRawData raw_data = ParseRawTable(current_sheet, ref table_range, filename, hash);
                                _table_raw_data_store.RawDataList.Add(raw_data);

                                Trace.WriteLine($"Table<{raw_data.TableName}> SR:{table_range.start_row} SC:{table_range.start_column} ER:{table_range.end_row} EC:{table_range.end_column}");
                            }
                        }
                    }
                }

                Marshal.FinalReleaseComObject(current_sheet);
            }

            // close
            Marshal.FinalReleaseComObject(worksheets);

            wb.Close(SaveChanges: false);
            Marshal.FinalReleaseComObject(wb);
        }

        private TableRange DetectTable(Excel.Worksheet worksheet, int table_start_row, int table_start_column)
        {
            var range = worksheet.UsedRange;

            int max_row = range.Rows.Count + 1;
            int max_column = range.Columns.Count + 1;
            bool detect_eot = false;
            bool detect_eof = false;

            TableRange table_range = new TableRange();
            table_range.start_row = table_start_row;
            table_range.start_column = table_start_column;

            double total_count = (max_column - table_start_column + 1) * (max_row - table_start_row + 2);
            for (int column = table_start_column; column <= max_column; ++column)
            {
                for (int row = table_start_row + 1; row <= max_row; ++row)
                {
                    var cell = worksheet.Cells[row, column];
                    if (cell.Value2 != null)
                    {
                        string cell_string = cell.Value2.ToString();
                        if (cell_string.Equals(LABEL_TABLE))
                        {
                            // 이건 언제든지 등장하면 안되는 녀석-_-
                            throw new ApplicationException("테이블이 종료되지 않은 채로 테이블이 또 시작");
                        }

                        if (cell_string.Equals(LABEL_EOT))
                        {
                            if (detect_eot)
                            {
                                throw new ApplicationException("테이블이 종료되었는데 또 테이블 종료가 발견");
                            }

                            // 테이블 종료를 발견했으니 다음 컬럼으로 진행
                            detect_eot = true;
                            table_range.end_row = row;
                            max_row = row;

                            break;
                        }

                        if (cell_string.Equals(LABEL_EOF))
                        {
                            if (detect_eof)
                            {
                                throw new ApplicationException("필드 종료 라벨 중복");
                            }

                            // 필드 종료를 발견했으니 완전히 종료
                            detect_eof = true;
                            table_range.end_column = column;

                            return table_range;
                        }
                    }
                }
            }

            throw new ApplicationException("테이블/필드 종료 라벨을 발견 못함");
            //return table_range;
        }

        private void Parse2(string filepath, string filename, string hash)
        {
            // open
            var wb = _excel_app.Workbooks.Open(filepath);
            var worksheets = wb.Worksheets;

            Trace.WriteLine($"Worksheets.Count={worksheets.Count}");

            // load all cells
            foreach (Excel.Worksheet current_sheet in worksheets)
            {
                if (current_sheet.Name.IndexOf('#') != -1)
                {
                    continue;
                }

                var range = current_sheet.UsedRange;
                Trace.WriteLine($"ActiveSheet.UsedRange Rows={range.Rows.Count}, Columns={range.Columns.Count} <{current_sheet.Name}>");

                int row_count = range.Rows.Count;
                int column_count = range.Columns.Count;

                if (row_count > 1000000 || column_count > 1000000)
                {
                    throw new ApplicationException($"COUNT OVERFLOW ; ActiveSheet.UsedRange Rows={range.Rows.Count}, Columns={range.Columns.Count} <{current_sheet.Name}>");
                }

                var start_cell = (Excel.Range)current_sheet.Cells[1, 1];
                var end_cell = (Excel.Range)current_sheet.Cells[row_count + 1, column_count + 1];
                var range_cell = current_sheet.get_Range(start_cell, end_cell);
                var range_values = (object[,])range_cell.Value2;

                for (int i = 1; i <= row_count; ++i)
                {
                    for (int j = 1; j <= column_count; ++j)
                    {
                        var cell = range_values[i, j];
                        if (cell != null)
                        {
                            string cell_string = Convert.ToString(cell);

                            if (cell_string.Equals(LABEL_TABLE))
                            {
                                // 테이블 디텍트
                                TableRange table_range = DetectTable2(current_sheet, i, j);

                                // raw data 생성
                                TableRawData raw_data = ParseRawTable(current_sheet, ref table_range, filename, hash);
                                _table_raw_data_store.RawDataList.Add(raw_data);

                                Trace.WriteLine($"Table<{raw_data.TableName}> SR:{table_range.start_row} SC:{table_range.start_column} ER:{table_range.end_row} EC:{table_range.end_column}");
                            }
                        }
                    }
                }

                Marshal.FinalReleaseComObject(current_sheet);
            }

            // close
            Marshal.FinalReleaseComObject(worksheets);

            wb.Close(SaveChanges: false);
            Marshal.FinalReleaseComObject(wb);
        }

        private TableRange DetectTable2(Excel.Worksheet worksheet, int table_start_row, int table_start_column)
        {
            var range = worksheet.UsedRange;

            int max_row = range.Rows.Count + 1;
            int max_column = range.Columns.Count + 1;
            bool detect_eot = false;
            bool detect_eof = false;

            TableRange table_range = new TableRange();
            table_range.start_row = table_start_row;
            table_range.start_column = table_start_column;

            var start_cell = (Excel.Range)worksheet.Cells[table_start_row + 1, table_start_column];
            var end_cell = (Excel.Range)worksheet.Cells[max_row, max_column];
            var range_cell = worksheet.get_Range(start_cell, end_cell);
            var range_values = (object[,])range_cell.Value2;

            for (int column = table_start_column; column <= max_column; ++column)
            {
                for (int row = table_start_row + 1; row <= max_row; ++row)
                {
                    int cell_row = row - table_start_row;
                    int cell_col = column - table_start_column + 1;
                    var cell = range_values[cell_row, cell_col];
                    if (cell != null)
                    {
                        string cell_string = Convert.ToString(cell);
                        if (cell_string.Equals(LABEL_TABLE))
                        {
                            // 이건 언제든지 등장하면 안되는 녀석-_-
                            throw new ApplicationException("테이블이 종료되지 않은 채로 테이블이 또 시작");
                        }

                        if (cell_string.Equals(LABEL_EOT))
                        {
                            if (detect_eot)
                            {
                                throw new ApplicationException("테이블이 종료되었는데 또 테이블 종료가 발견");
                            }

                            // 테이블 종료를 발견했으니 다음 컬럼으로 진행
                            detect_eot = true;
                            table_range.end_row = row;
                            max_row = row;

                            break;
                        }

                        if (cell_string.Equals(LABEL_EOF))
                        {
                            if (detect_eof)
                            {
                                throw new ApplicationException("필드 종료 라벨 중복");
                            }

                            // 필드 종료를 발견했으니 완전히 종료
                            detect_eof = true;
                            table_range.end_column = column;

                            return table_range;
                        }
                    }
                }
            }

            throw new ApplicationException("테이블/필드 종료 라벨을 발견 못함");
            //return table_range;
        }

        private TableRawData ParseRawTable(Excel.Worksheet worksheet, ref TableRange range, string filename, string hash)
        {
            TableRawData raw_data = new TableRawData();

            if (worksheet.Cells[range.start_row, range.start_column + 1].Value2 == null)
            {
                throw new ApplicationException("테이블 이름이 없습니다");
            }

            // 테이블 이름
            raw_data.TableName = worksheet.Cells[range.start_row, range.start_column + 1].Value2.ToString().Trim();

            // 파일명 및 해시
            raw_data.FileName = filename;
            raw_data.Hash = hash;

            // 필드 이름/타입
            raw_data.FieldNames = new List<string>(range.end_column - range.start_column - 1);
            raw_data.FieldNameCounts = new List<TableRawFieldNameCount>(range.end_column - range.start_column - 1);
            raw_data.FieldTypes = new List<TableFieldType>(range.end_column - range.start_column - 1);
            raw_data.FieldTypeNames = new List<string>(range.end_column - range.start_column - 1);

            List<int> skipped_columns = new List<int>();

            for (int column = range.start_column + 1; column <= range.end_column - 1; ++column)
            {
                string field_name = worksheet.Cells[range.start_row + 1, column].Value2.ToString().Trim();
                if (field_name.IndexOf('#') != -1)
                {
                    skipped_columns.Add(column);
                    continue;
                }

                raw_data.FieldNames.Add(field_name);
                var field_name_counts = raw_data.FieldNameCounts.Find(x => x.FieldName == field_name);
                if (field_name_counts == null)
                {
                    field_name_counts = new TableRawFieldNameCount() { FieldName = field_name, Count = 0 };
                    raw_data.FieldNameCounts.Add(field_name_counts);
                }
                field_name_counts.Count += 1;

                if (worksheet.Cells[range.start_row + 2, column].Value2 == null)
                {
                    throw new ApplicationException(string.Format("필드 이름 없음 - Table={0}, FieldName={1}", raw_data.TableName, field_name));
                }

                string field_type_name = worksheet.Cells[range.start_row + 2, column].Value2.ToString().Trim();
                raw_data.FieldTypes.Add(GetFieldType(field_type_name));
                raw_data.FieldTypeNames.Add(field_type_name);
            }

            // 레코드
            raw_data.Records = new List<TableRawRecords>(range.end_row - range.start_row - 3);

            {
                var start_cell = (Excel.Range)worksheet.Cells[range.start_row + 3, range.start_column + 1];
                var end_cell = (Excel.Range)worksheet.Cells[range.end_row - 1, range.end_column - 1];
                var range_cell = worksheet.get_Range(start_cell, end_cell);
                var range_values = (object[,])range_cell.Value2;

                int row_count = range_values.GetLength(0);
                int col_count = range_values.GetLength(1);
                if (row_count == 0)
                {
                    throw new ApplicationException("테이블 내용이 없습니다: " + raw_data.TableName);
                }

                for (int row = 1; row <= row_count; ++row)
                {
                    string[] record_array = new string[raw_data.FieldNames.Count];

                    int index = 0;
                    for (int col = 1; col <= col_count; ++col)
                    {
                        if (skipped_columns.Contains(range.start_column + col))
                        {
                            //Trace.WriteLine("필드 넘김: {0} {1}", range.start_column + col, range_values[row, col]);
                            continue;
                        }

                        var cell = range_values[row, col];
                        if (cell != null)
                        {
                            record_array[index++] = Convert.ToString(cell);
                        }
                        else
                        {
                            record_array[index++] = null;
                        }
                    }

                    raw_data.Records.Add(new TableRawRecords(record_array));
                }
            }

            return raw_data;
        }

        private TableFieldType GetFieldType(string field_type_string)
        {
            int exc_idx = field_type_string.IndexOf('!');
            TableFieldType table_field_type = TableFieldType.INVALID;

            if (exc_idx == 0)
            {
                if (!Enum.TryParse<TableFieldType>(field_type_string.Substring(1), out table_field_type))
                {
                    if (field_type_string.Substring(1, 2) == TableFieldType.ID.ToString())
                    {
                        return TableFieldType.ID;
                    }
                    else if (field_type_string.Substring(1, 7) == TableFieldType.ENUM_ID.ToString())
                    {
                        return TableFieldType.ENUM_ID;
                    }
                    else if (field_type_string.Substring(1, 8) == TableFieldType.KEY_ENUM.ToString())
                    {
                        return TableFieldType.KEY_ENUM;
                    }
                    else
                    {
                        throw new ApplicationException("필드 타입 문자열 오류: " + field_type_string);
                    }
                }
            }
            else
            {
                throw new ApplicationException("필드 타입 문자열 오류: " + field_type_string);
            }

            return table_field_type;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    // dispose managed resource
                    if (_excel_app != null)
                    {
                        _excel_app.Quit();
                        Marshal.FinalReleaseComObject(_excel_app);

                        _excel_app = null;
                        GC.Collect();
                    }
                }

                // dispose unmanaged resource
                _disposed = true;
            }
        }

        ~ExcelTableBuild()
        {
            Dispose(false);
        }
    }
}