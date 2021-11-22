using EMDD.Reporting.Line;
using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    /// <summary>
    /// Create a Graph
    /// </summary>
    public class LineGraph : LineContent
    {
        /// <summary>
        /// Graph
        /// </summary>
        public LineGraph(uint tabLevel) :base(tabLevel)
        {
#pragma warning disable RCS1079 // Throwing of new NotImplementedException.
            throw new System.NotImplementedException("Still working on LineGraphs");
#pragma warning restore RCS1079 // Throwing of new NotImplementedException.
        }

        internal override void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
        }
    }
}