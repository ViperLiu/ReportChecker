using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker.TableParsers
{
    class TestInfoTableParser : TableParser
    {
        public int AcceptCount { get; private set; }

        public int FailCount { get; private set; }

        public int NotfitCount { get; private set; }

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
