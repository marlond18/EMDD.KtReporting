using EMDD.Reporting.Line;

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
        internal override void WriteToString(ref StringBuilder str)
        {
            str.Append(new string('\t', (int)TabIndex)).AppendLine("<Graph Not converted To basic String>");
        }
    }
}