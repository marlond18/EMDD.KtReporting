using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    namespace Line
    {
        /// <summary>
        /// Abstract Class for Line of Texts
        /// </summary>
        public abstract class LineContent
        {
            protected LineContent(uint tabLevel)
            {
                _tabIndex = tabLevel;
            }

            internal readonly uint _tabIndex;
            //internal abstract void WriteLineOpenXML(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0);

            internal abstract void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0);

            internal void CreateLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
            {
                WriteLine(range, justify, fontsize, leftIndent * (int)_tabIndex, spaceAfter, bold);
                range.InsertParagraphAfter();
            }
        }
    }
}