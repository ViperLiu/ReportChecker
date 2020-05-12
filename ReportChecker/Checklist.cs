using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportChecker
{
    class Checklist
    {
        private List<ITarget> _rulesToBeChecked = new List<ITarget>();

        public void AddRule<T>() where T : ITarget, new()
        {
            T rule = new T();
            _rulesToBeChecked.Add(rule);
        }

        public void CheckAll()
        {
            foreach(var target in _rulesToBeChecked)
            {
                target.Check();
            }
        }
    }
}
