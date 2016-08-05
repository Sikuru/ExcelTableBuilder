using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sikuru.ExcelTableBuilder
{
    class Program
    {
        static int Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));

            string bin_filename = null;
            string namespace_name = null;

            if (args == null || args.Length < 1)
            {
                Trace.WriteLine("ExcelTableBuilder [생성될 소스파일의 네임스페이스] {바이너리 파일명}");
            }
            else
            {
                namespace_name = args[0].Trim();

                if (args.Length >= 2)
                {
                    bin_filename = args[1].Trim();
                }
                else
                {
                    // $"{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.csv"
                    bin_filename = "";
                }
            }

            Process[] excel_proc = Process.GetProcessesByName("EXCEL");
            if (excel_proc.Length > 0)
            {
                Trace.WriteLine("동작중인 엑셀 프로세스가 있습니다. 열려있는 모든 엑셀을 종료하세요. 종료해도 오류가 난다면, 작업관리자에서 EXCEL.EXE 를 모두 종료 후 재시도 해 주세요.");
                Environment.Exit(1);
                return 1;
            }

            if (string.IsNullOrEmpty(namespace_name))
            {
                Trace.WriteLine("ExcelTableBuilder [생성될 소스파일의 네임스페이스] {바이너리 파일명}");
                Environment.Exit(1);
                return 1;
            }

            try
            {
                var builder = new Builder();
                builder.Build(Directory.GetCurrentDirectory(), bin_filename, namespace_name);
                Trace.WriteLine("빌드 완료.");
                Environment.Exit(0);
                return 0;
            }
            catch (Exception e)
            {
                Trace.WriteLine("빌드 오류: {0}", e.Message);
                Environment.Exit(1);
                return 1;
            }
        }
    }
}
