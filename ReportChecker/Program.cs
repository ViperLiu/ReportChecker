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
        static void Main(string[] args)
        {
            string fileName;
            if(args.Length < 1)
            {
                Console.Write("請選擇檔案：");
                fileName = Console.ReadLine();
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
            var classText = InputDocx.Tables[0].Rows[8].Cells[1].Paragraphs.First().Text;
            var acceptCount = int.Parse(InputDocx.Tables[1].Rows[10].Cells[1].Paragraphs.First().Text);
            var failCount = int.Parse(InputDocx.Tables[1].Rows[10].Cells[2].Paragraphs.First().Text);
            var notfitCount = int.Parse(InputDocx.Tables[1].Rows[10].Cells[3].Paragraphs.First().Text);
            var totalCount = acceptCount + failCount + notfitCount;

            //Console.WriteLine("= = = = = APP資訊表 = = = = =");
            //Console.WriteLine("檢測分類：{0}", classText);
            //Console.WriteLine("符合數量：{0}", acceptCount);
            //Console.WriteLine("不符合數量：{0}", failCount);
            //Console.WriteLine("不適用數量：{0}", notfitCount);
            //Console.WriteLine("總共數量：{0}", totalCount);
            //Console.WriteLine();

            var overviewTable = InputDocx.Tables[2];
            var overviewTableResults = CountOverviewTable(overviewTable);
            //Console.WriteLine("= = = = = 檢測結果摘要表 = = = = =");
            //Console.WriteLine("符合數量：{0}", overviewTableResults[0]);
            //Console.WriteLine("不符合數量：{0}", overviewTableResults[1]);
            //Console.WriteLine("不適用數量：{0}", overviewTableResults[2]);
            //Console.WriteLine("總共數量：{0}", overviewTableResults[3]);
            //Console.WriteLine();

            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜　　　｜資訊表｜摘要表｜");
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜符合　｜　{0,2:##}　｜　{1,2:##}　｜", acceptCount, overviewTableResults[0]);
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜不符合｜　{0,2:##}　｜　{1,2:##}　｜", failCount, overviewTableResults[1]);
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine("｜不適用｜　{0,2:##}　｜　{1,2:##}　｜", notfitCount, overviewTableResults[2]);
            Console.WriteLine("－－－－－－－－－－－－－");
            Console.WriteLine();

            GetAllRuleNumbers();
            //Console.WriteLine(RuleNumberList.Count);

            GetAllResultTables();
            //Console.WriteLine(ResultTableList.Count);

            MergeRuleNumberAndResultTables();
            foreach(var target in TargetsToCheck)
            {
                target.Check();
                FullErrorsList.AddRange(target.Errors);
                target.PrintErrors();
            }
            //foreach(var err in FullErrorsList)
            //{
            //    Console.WriteLine(err);
            //}
            //foreach(var item in inputDocx.Lists[1].Items)
            //{
            //    Console.WriteLine(item.StyleName);
            //    Console.WriteLine(item.Text.Trim());
            //}
            Console.WriteLine("掃描完成");
            Console.ReadLine();
        }

        private static int[] CountOverviewTable(Table tab)
        {
            int[] results = new int[4];
            for(var i = 1; i < tab.RowCount; i++)
            {
                var text = tab.Rows[i].Cells[2].Paragraphs[0].Text;
                if (text == Strings.AcceptText)
                    results[0]++;
                else if (text == Strings.FailText)
                    results[1]++;
                else if (text == Strings.NotfitText)
                    results[2]++;
                else
                    Console.WriteLine(new ReportError("檢測結果摘要表", -1, "文字錯誤"));
            }
            results[3] = results[0] + results[1] + results[2];
            return results;
        }

        private static void MergeRuleNumberAndResultTables()
        {
            if (RuleNumberList.Count != ResultTableList.Count)
                return;

            for(var i = 0; i < RuleNumberList.Count; i++)
            {
                TargetsToCheck.Add(new RuleResultTarget(RuleNumberList[i], ResultTableList[i]));
            }
        }

        private static void GetAllRuleNumbers()
        {
            foreach (var p in InputDocx.Paragraphs)
            {
                if (p.StyleName == "4-11" && !p.IsListItem)
                {
                    var ruleNumber = p.Text.Trim().Substring(0, 9);
                    RuleNumberList.Add(ruleNumber);
                    //Console.WriteLine(ruleNumber);
                }
            }
        }

        private static void GetAllResultTables()
        {
            foreach(var tab in InputDocx.Tables)
            {
                if(tab.Rows[0].Cells[0].Paragraphs[0].Text == "檢測基準")
                {
                    ResultTableList.Add(tab);
                }
            }
        }

        private void CheckResultCount()
        {

        }
    }
}
