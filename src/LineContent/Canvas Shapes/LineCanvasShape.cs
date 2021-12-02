
namespace EMDD.Reporting
{
    /// <summary>
    /// straight line
    /// </summary>
    public class LineCanvasShape : CanvasShapes
    {
        /// <summary>
        /// intialize
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="thickness"></param>
        public LineCanvasShape((double x, double y) start, (double x, double y) end, uint tabLevel, double thickness = 1) :base(tabLevel)
        {
            Start = start;
            End = end;
            Thickness = thickness;
        }

        public (double x, double y) Start { get; }
        public (double x, double y) End { get; }
        public  double Thickness { get; }
    }
}