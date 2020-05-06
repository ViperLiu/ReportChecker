using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportChecker
{
    class ReportError
    {
        public string Description { get; private set; }
        public string RuleNumber { get; private set; }
        public int SubRuleNumber { get; private set; }

        public ReportError(string ruleNumber, int subRuleNumber, string description)
        {
            RuleNumber = ruleNumber;
            SubRuleNumber = subRuleNumber;
            Description = description;
        }

        public override string ToString()
        {
            if(SubRuleNumber == -1)
                return string.Format("{0}, {2}", RuleNumber, SubRuleNumber, Description);
            return string.Format("{0}, 基準({1}), {2}", RuleNumber, SubRuleNumber, Description);
        }
    }
}
