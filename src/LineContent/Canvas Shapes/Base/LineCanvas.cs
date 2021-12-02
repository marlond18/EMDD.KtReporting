using EMDD.Reporting.Line;

using System.Text;

namespace EMDD.Reporting
{
    /// <summary>
    /// create shape containers
    /// </summary>
    public class LineCanvas : LineContent
    {
        public (double X, double Y) Location { get; }
        public (double Width, double Height) Size { get; }
        public List<CanvasShapes> Shapes { get; }

        /// <summary>
        /// Initialize location and size
        /// </summary>
        /// <param name="topLeftX"></param>
        /// <param name="topLeftY"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public LineCanvas(double topLeftX, double topLeftY, double width, double height, uint tabLevel): base(tabLevel)
        {
            Location = (topLeftX, topLeftY);
            Size = (width, height);
            Shapes = new List<CanvasShapes>();
        }

        /// <summary>
        /// add additional shapes to the drawing
        /// </summary>
        /// <param name="shape"></param>
        public void AddShape(CanvasShapes shape)
        {
            Shapes.Add(shape);
        }

        /// <summary>
        /// directly addshape line to the doc file
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="thickness"></param>
        public void AddLine((double x, double y) start, (double x, double y) end, double thickness = 1)
        {
            AddShape(new LineCanvasShape(start, end, TabIndex, thickness));
        }

        /// <summary>
        /// add curve directly to the canvas
        /// </summary>
        /// <param name="points"></param>
        public void AddCurve(params (double x, double y)[] points)
        {
            AddShape(new CurveCanvasShape(1, TabIndex, points));
        }

        internal override void WriteToString(ref StringBuilder str)
        {
            str.Append(new string('\t', (int)TabIndex)).AppendLine("<Canvas Not converted To basic String>");
        }
    }
}