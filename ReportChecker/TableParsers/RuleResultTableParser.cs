using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker.TableParsers
{
    class RuleResultTableParser : TableParser
    {
        private int _subRuleCount;
        public RuleResultTableParser(Table table) : base(table)
        {
            _subRuleCount = _table.RowCount - 5;
        }

        public Paragraph GetSubRuleResult(int subRuleNumber)
        {
            if (subRuleNumber > _subRuleCount)
                return null;

            return GetCell(subRuleNumber, 0).Paragraphs.First();
        }

        public ICollection<Paragraph> GetTestResults()
        {
            return GetCell(_table.RowCount - 1, 0).Paragraphs;
        }

        public int GetSubRuleCount()
        {
            return _subRuleCount;
        }
    }
}
