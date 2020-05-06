using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace ReportChecker
{
    class RuleResultTarget : ITarget
    {
        public readonly int SubRuleCount;

        public string Title { get; private set; }

        public Table ResultTable { get; private set; }

        public List<ReportError> Errors { get; }

        public RuleResultTarget(string ruleNumber, Table resultTable)
        {
            Title = ruleNumber;
            ResultTable = resultTable;
            Errors = new List<ReportError>();
            SubRuleCount = ResultTable.RowCount - 5;
        }

        public void PrintErrors()
        {
            if(Errors.Count == 0)
            {
                return;
            }
            Console.WriteLine("= = = = = {0} = = = = =", Title);
            foreach(var err in Errors)
            {
                Console.WriteLine(err);
            }
            Console.WriteLine();
        }

        public void Check()
        {
            _checkUpperTextStyle();
            _checkResultResultTextStyle();
        }

        private void _checkUpperTextStyle()
        {
            for(var i = 1; i <= SubRuleCount; i++)
            {
                if(ResultTable.Rows[i].Cells[0].Paragraphs[0].MagicText.Count == 0)
                {
                    Errors.Add(new ReportError(Title, i, "上方欄位空白"));
                    continue;
                }
                var text = ResultTable.Rows[i].Cells[0].Paragraphs[0].Text;
                var formatting = ResultTable.Rows[i].Cells[0].Paragraphs[0].MagicText[0].formatting;
                var color = formatting.FontColor == null ? Color.Black : (Color)(formatting.FontColor);
                var fontSize = formatting.Size == null ? 12 : formatting.Size;
                //Console.WriteLine("Text={0}, Color={1}, Size={2}", text, color, fontSize);

                if (fontSize != 12)
                    Errors.Add(new ReportError(Title, i, "上方字體大小錯誤"));

                var condition1 = (text == Strings.AcceptText && color.ToArgb() != Color.Black.ToArgb());
                var condition2 = (text == Strings.FailText && color.ToArgb() != Color.Red.ToArgb());
                var condition3 = (text == Strings.NotfitText && color.ToArgb() != Color.Blue.ToArgb());
                if (condition1 || condition2 || condition3)
                {
                    Console.WriteLine("{0}, {1}", text, color);
                    Errors.Add(new ReportError(Title, i, "上方文字顏色錯誤"));
                }
                else if(text != Strings.AcceptText && text != Strings.FailText && text != Strings.NotfitText)
                {
                    Errors.Add(new ReportError(Title, i, "上方文字錯誤"));
                }
            }
        }

        private void _checkResultResultTextStyle()
        {
            foreach(var p in ResultTable.Rows.Last().Paragraphs)
            {
                //取得子基準的編號
                var tmpArray = p.Text.Split(new string[] { "基準(" }, StringSplitOptions.RemoveEmptyEntries);
                int subRuleNumber;
                if (tmpArray.Length <= 1)
                {
                    subRuleNumber = -1;
                }
                else
                {
                    tmpArray = tmpArray[1].Split(
                        new string[] { ")" },
                        StringSplitOptions.RemoveEmptyEntries
                        );
                    subRuleNumber = int.Parse(tmpArray[0]);
                }

                //檢查文字描述是否與上方欄位一致
                _checkConsistency(subRuleNumber, p.Text);

                //檢查文字大小是否有誤
                foreach (var t in p.MagicText)
                {
                    var fontSize = t.formatting.Size == null ? 12 : t.formatting.Size;
                    if (fontSize != 12)
                    {
                        Errors.Add(new ReportError(Title, subRuleNumber, "下方文字描述的字體大小錯誤"));
                    }
                }

                //檢查文字格式是否有誤
                var pattern = string.Format("符合基準({0})。", subRuleNumber);
                if(p.Text.TrimEnd().EndsWith(pattern))
                {
                    Errors.Add(new ReportError(Title, subRuleNumber, "下方文字描述錯誤，應將 符合/不符合 放後面"));
                }
            }
        }

        private void _checkConsistency(int subRuleNumber, string resultText)
        {
            if (subRuleNumber == -1)
                return;
            var targetText = ResultTable.Rows[subRuleNumber].Paragraphs[0].Text;
            var condition1 = (targetText == Strings.FailText && !resultText.Contains(Strings.FailText));
            var condition2 = (targetText == Strings.NotfitText && !resultText.Contains(Strings.NotfitText));
            var condition3 = (targetText == Strings.AcceptText && (resultText.Contains(Strings.FailText) || resultText.Contains(Strings.NotfitText)));
            if (condition1 || condition2 || condition3)
            {
                Errors.Add(new ReportError(Title, subRuleNumber, "下方文字描述與上方判定不符"));
            }
        }
    }
}
