using ReportChecker.TableParsers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ReportChecker
{
    class Program
    {
        private static List<string> RuleNumberList = new List<string>();
        private static List<Table> ResultTableList = new List<Table>();
        private static List<ITarget> TargetsToCheck = new List<ITarget>();
        private static List<ReportError> FullErrorsList = new List<ReportError>();
        private static DocX InputDocx;
        private static TableParserProvider tableParsers;
        private static ReportCoverTableParser ReportCoverTable;
        private static IResultCountable TestInfoTable;
        private static IResultCountable OverviewTable;

        static void Main(string[] args)
        {
            string fileName;
            if(args.Length < 1)
            {
                Console.Write("請選擇檔案：");
                fileName = Console.ReadLine();
                Console.WriteLine();
            }
            else if(args.Length == 1)
            {
                fileName = args[0];
            }
            else
            {
                Console.WriteLine("參數錯誤");
                return;
            }
            InputDocx = DocX.Load(fileName);

            tableParsers = new TableParserProvider(InputDocx);
            ReportCoverTable = tableParsers.GetReportCoverTableParser();
            TestInfoTable = tableParsers.GetTestInfoTableParser();
            OverviewTable = tableParsers.GetOverviewTableParser();

            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜　　　｜資訊表｜摘要表｜");
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜符合　｜　{0,2:##}　｜　{1,2:##}　｜", TestInfoTable.AcceptCount, OverviewTable.AcceptCount);
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜不符合｜　{0,2:##}　｜　{1,2:##}　｜", TestInfoTable.FailCount, OverviewTable.FailCount);
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜不適用｜　{0,2:##}　｜　{1,2:##}　｜", TestInfoTable.NotfitCount, OverviewTable.NotfitCount);
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine();

            Console.WriteLine("掃描完成");
            Console.ReadLine();
        }
    }
}
