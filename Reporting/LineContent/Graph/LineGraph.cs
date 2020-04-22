using Reporting.Line;
using System;
using Microsoft.Office.Interop.Word;

namespace Reporting
{
    /// <summary>
    /// Create a Graph
    /// </summary>
    public class LineGraph : LineContent
    {
        /// <summary>
        /// Graph
        /// </summary>
        public LineGraph() => throw new NotImplementedException("Still working on LineGraphs");

        internal override void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {

        }
    }
}