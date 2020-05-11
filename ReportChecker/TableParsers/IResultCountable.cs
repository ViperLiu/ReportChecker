using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportChecker.TableParsers
{
    interface IResultCountable
    {
        int AcceptCount { get; }
        int FailCount { get; }
        int NotfitCount { get; }
    }
}
