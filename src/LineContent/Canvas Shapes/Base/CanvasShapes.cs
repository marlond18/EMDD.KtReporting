using EMDD.Reporting.Line;

namespace EMDD.Reporting
{
    /// <summary>
    /// base shape for canvas
    /// </summary>
    public abstract class CanvasShapes : LineContent
    {
        protected CanvasShapes(uint tabLevel) : base(tabLevel)
        {
        }

        internal override void WriteToString(ref StringBuilder str)
        {
            str.Append(new string('\t', (int)TabIndex)).AppendLine("<Shape Not converted To basic String>");
        }
    }
}