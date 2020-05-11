using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker.TableParsers
{
    class TestInfoTableParser : TableParser, IResultCountable
    {
        public int AcceptCount { get; private set; } = 0;

        public int FailCount { get; private set; } = 0;

        public int NotfitCount { get; private set; } = 0;

        public TestInfoTableParser(Table table)
                : base(table)
        {
            GetAcceptCount();
            GetFailCount();
            GetNotfitCount();
        }

        private void GetAcceptCount()
        {
            var cell = GetCell(10, 1);
            AcceptCount = int.Parse(cell.Paragraphs.First().Text);
        }

        private void GetFailCount()
        {
            var cell = GetCell(10, 2);
            FailCount = int.Parse(cell.Paragraphs.First().Text);
        }

        private void GetNotfitCount()
        {
            var cell = GetCell(10, 3);
            NotfitCount = int.Parse(cell.Paragraphs.First().Text);
        }
    }
}
