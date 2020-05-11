using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker.TableParsers
{
    class OverviewTableParser : TableParser
    {
        public int AcceptCount { get; private set; } = 0;

        public int FailCount { get; private set; } = 0;

        public int NotfitCount { get; private set; } = 0;

        public Dictionary<string, Paragraph> OverviewTableDict { get; } 
            = new Dictionary<string, Paragraph>();

        public OverviewTableParser(Table table) : base(table)
        {
            GetTestResultCounts();
            BuildOverviewTableDictionary();
        }

        private void GetTestResultCounts()
        {
            for(var i = 1; i < _table.RowCount; i++)
            {
                var text = GetCell(i, 2).Paragraphs.First().Text;
                if (text == Strings.AcceptText)
                    AcceptCount++;
                else if (text == Strings.FailText)
                    FailCount++;
                else if (text == Strings.NotfitText)
                    NotfitCount++;
            }
        }

        private void BuildOverviewTableDictionary()
        {
            for (var i = 1; i < _table.RowCount; i++)
            {
                var paragraph = GetCell(i, 2).Paragraphs.First();
                var ruleNumber = GetCell(i, 1).Paragraphs.First().Text.Substring(0, 9);

                OverviewTableDict.Add(ruleNumber, paragraph);
            }
        }

        public Paragraph GetParagraph(string rulenumber)
        {
            OverviewTableDict.TryGetValue(rulenumber, out Paragraph para);
            return para;
        }
    }
}
