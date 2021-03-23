using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    /// <summary>
    /// straight line
    /// </summary>
    public class LineCanvasShape : CanvasShapes
    {
        internal override void DrawShapeOnDoc(Document doc)
        {
            var line = doc.Shapes.AddLine((float)_start.x, (float)_start.y, (float)_end.x, (float)_end.y);
            line.Line.Weight = (float)_thickness;
        }

        internal override void DrawShapeOnCanvas(Word.CanvasShapes canvasItems)
        {
            var line = canvasItems.AddLine((float)_start.x, (float)_start.y, (float)_end.x, (float)_end.y);
            line.Line.Weight = (float)_thickness;
        }

        /// <summary>
        /// intialize
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="thickness"></param>
        public LineCanvasShape((double x, double y) start, (double x, double y) end, double thickness = 1)
        {
            _start = start;
            _end = end;
            _thickness = thickness;
        }

        private (double x, double y) _start;
        private (double x, double y) _end;
        private readonly double _thickness;
    }
}