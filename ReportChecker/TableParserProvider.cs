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
    class TableParserProvider
    {
        private readonly Table _reportCoverTable;
        private readonly Table _testInformationTable;
        private readonly Table _overviewTable;
        

        private readonly DocX _targetDocument;
        private readonly TableParser _testInfoTableParser;
        private readonly TableParser _reportCoverTableParser;
        private readonly Dictionary<string, RuleResultTableParser> _ruleResultTableParsers 
            = new Dictionary<string, RuleResultTableParser>();

        public TableParserProvider(DocX targetDocument)
        {
            _targetDocument = targetDocument;
            GetAllRuleResultTables();
            foreach (var table in targetDocument.Tables)
            {
                var topLeftCell = table.Rows[0].Cells[0];
                if(topLeftCell.Paragraphs[0].Text == "檢測公司名稱")
                {
                    _reportCoverTableParser = new ReportCoverTableParser(table);
                }
                else if(topLeftCell.Paragraphs[0].Text == "送測單位")
                {
                    _testInfoTableParser = new TestInfoTableParser(table);
                }
                else if(topLeftCell.Paragraphs[0].Text == "#")
                {
                    _overviewTable = table;
                }
            }
        }

        public TableParser GetReportCoverTableParser()
        {
            return _reportCoverTableParser;
        }

        public TableParser GetTestInfoTableParser()
        {
            return _testInfoTableParser;
        }

        public TableParser GetRuleResultTableParser(string ruleNumber)
        {
            if (!_ruleResultTableParsers.ContainsKey(ruleNumber))
                return null;
            return _ruleResultTableParsers[ruleNumber];
        }

        private void GetAllRuleResultTables()
        {
            var ruleNumberList = new List<string>();
            var ruleResultTables = new List<Table>();
            foreach (var p in _targetDocument.Paragraphs)
            {
                if (p.StyleName == "4-11" && !p.IsListItem)
                {
                    var ruleNumber = p.Text.Trim().Substring(0, 9);
                    ruleNumberList.Add(ruleNumber);
                }
            }
            foreach(var table in _targetDocument.Tables)
            {
                var topLeftCell = table.Rows[0].Cells[0];
                if (topLeftCell.Paragraphs[0].Text == "檢測基準")
                {
                    ruleResultTables.Add(table);
                }
            }

            if (ruleNumberList.Count != ruleResultTables.Count)
                throw new Exception("Report parsing error");

            for (var i = 0; i < ruleNumberList.Count; i++)
            {
                _ruleResultTableParsers.Add(ruleNumberList[i], new RuleResultTableParser(ruleResultTables[i]));
            }
        }
    }
}
