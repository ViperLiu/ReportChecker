using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker.TableParsers
{
    class ReportCoverTableParser : TableParser
    {
        public string TestClass { get; private set; }
        public string AppName { get; private set; }
        public string AppPlatform { get; private set; }
        public string ReportId { get; private set; }
        public string CaseDate { get; private set; }
        public string ReportDate { get; private set; }
        public string ClientCompanyName { get; private set; }
        public string ClientCompanyAddress { get; private set; }

        public ReportCoverTableParser(Table table) : base(table)
        {
            ParseData();
        }

        private void ParseData()
        {
            foreach(var row in _table.Rows)
            {
                var rowName = row.Cells[0].Paragraphs.First().Text;
                var rowValue = row.Cells[1].Paragraphs.First().Text;
                switch (rowName)
                {
                    case "送檢單位名稱":
                        ClientCompanyName = rowValue;
                        break;
                    case "送檢單位地址":
                        ClientCompanyAddress = rowValue;
                        break;
                    case "送檢APP名稱":
                        AppName = rowValue;
                        break;
                    case "APP作業系統":
                        AppPlatform = rowValue;
                        break;
                    case "檢測分類":
                        TestClass = rowValue;
                        break;
                    case "報告編號":
                        ReportId = rowValue;
                        break;
                    case "收件日期":
                        CaseDate = rowValue;
                        break;
                    case "報告日期":
                        ReportDate = rowValue;
                        break;
                }
            }
        }
    }
}
