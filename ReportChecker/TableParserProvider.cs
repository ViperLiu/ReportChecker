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
        private readonly List<Table> _ruleResultTables = new List<Table>();

        private readonly DocX _targetDocument;
        private readonly TableParser _testInfoTableParser;

        public TableParserProvider(DocX targetDocument)
        {
            _targetDocument = targetDocument;
            foreach(var table in targetDocument.Tables)
            {
                var topLeftCell = table.Rows[0].Cells[0];
                if(topLeftCell.Paragraphs[0].Text == "檢測公司名稱")
                {
                    _reportCoverTable = table;
                }
                else if(topLeftCell.Paragraphs[0].Text == "送測單位")
                {
                    _testInfoTableParser = new TestInfoTableParser(table);
                }
                else if(topLeftCell.Paragraphs[0].Text == "#")
                {
                    _overviewTable = table;
                }
                else if (topLeftCell.Paragraphs[0].Text == "檢測基準")
                {
                    _ruleResultTables.Add(table);
                }
            }
        }
    }
}
