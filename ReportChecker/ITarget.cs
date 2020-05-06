using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker
{
    interface ITarget
    {
        string Title { get; }

        Table ResultTable { get; }

        List<ReportError> Errors { get; }

        void Check();

        void PrintErrors();
    }
}
