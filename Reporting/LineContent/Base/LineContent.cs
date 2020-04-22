using Microsoft.Office.Interop.Word;

namespace Reporting
{
    namespace Line
    {
        /// <summary>
        /// Abstract Class for Lines
        /// </summary>
        public abstract class LineContent
        {
            //internal abstract void WriteLineOpenXML(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0);

            internal abstract void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0);

            internal void CreateLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
            {
                WriteLine(range, justify, fontsize, leftIndent, spaceAfter, bold);
                range.InsertParagraphAfter();
            }
        }
    }
}