using EMDD.Reporting.Line;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    /// <summary>
    /// create shape containers
    /// </summary>
    public class LineCanvas : LineContent
    {
        private (double X, double Y) _location;
        private (double Width, double Height) _size;
        private readonly List<CanvasShapes> _shapes;

        /// <summary>
        /// Initialize location and size
        /// </summary>
        /// <param name="topLeftX"></param>
        /// <param name="topLeftY"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public LineCanvas(double topLeftX, double topLeftY, double width, double height)
        {
            _location = (topLeftX, topLeftY);
            _size = (width, height);
            _shapes = new List<CanvasShapes>();
        }

        /// <summary>
        /// add additional shapes to the drawing
        /// </summary>
        /// <param name="shape"></param>
        public void AddShape(CanvasShapes shape)
        {
            _shapes.Add(shape);
        }

        /// <summary>
        /// directly addshape line to the doc file
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="thickness"></param>
        public void AddLine((double x, double y) start, (double x, double y) end, double thickness = 1)
        {
            AddShape(new LineCanvasShape(start, end, thickness));
        }

        /// <summary>
        /// add curve directly to the canvas
        /// </summary>
        /// <param name="points"></param>
        public void AddCurve(params (double x, double y)[] points)
        {
            AddShape(new CurveCanvasShape(1, points));
        }

        internal override void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
            var canvas = range.Document.Shapes.AddCanvas((float)_location.X, (float)_location.Y, (float)_size.Width, (float)_size.Height, range);
            canvas.WrapFormat.Type = WdWrapType.wdWrapInline;
            foreach (var shape in _shapes)
            {
                shape.DrawShapeOnCanvas(canvas.CanvasItems);
            }
            canvas.CanvasItems.SelectAll();
            range.Application.Selection.Cut();
            canvas.Delete();
            range.PasteSpecial(DataType: WdPasteDataType.wdPasteEnhancedMetafile);
        }
    }
}