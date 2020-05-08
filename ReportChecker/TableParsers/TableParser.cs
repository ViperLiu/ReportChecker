using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker.TableParsers
{
    class TableParser
    {
        protected readonly Table _table;

        public TableParser(Table table)
        {
            _table = table;
        }

        public Cell GetCell(int row, int col)
        {
            return _table.Rows[row].Cells[col];
        }
    }
}
