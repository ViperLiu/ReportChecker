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
        public TestInfoTableParser(Table table)
                : base(table)
        {
        }

        public int GetAcceptCount()
        {
            var cell = GetCell(10, 1);
            var acceptCount = int.Parse(cell.Paragraphs.First().Text);
            return acceptCount;
        }

        public int GetFailCount()
        {
            var cell = GetCell(10, 2);
            var failCount = int.Parse(cell.Paragraphs.First().Text);
            return failCount;
        }

        public int GetNotfitCount()
        {
            var cell = GetCell(10, 3);
            var notfitCount = int.Parse(cell.Paragraphs.First().Text);
            return notfitCount;
        }
    }
}
