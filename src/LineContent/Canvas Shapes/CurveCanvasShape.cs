using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System;

namespace EMDD.Reporting
{
    /// <summary>
    /// Draw a curve line
    /// </summary>
    public class CurveCanvasShape : CanvasShapes
    {
        private readonly double _thickness;
        private readonly (double x, double y)[] _points;

        /// <summary>
        /// initialize the points needed
        /// </summary>
        /// <param name="thickness"></param>
        /// <param name="points"></param>
        public CurveCanvasShape(double thickness, uint tabLevel, params (double x, double y)[] points) : base(tabLevel)
        {
            _points = points;
            _thickness = thickness;
        }

        internal override void DrawShapeOnCanvas(Word.CanvasShapes canvasItems)
        {
            canvasItems.AddCurve(ConvertTuplePointsToSafePoints()).Line.Weight = (float)_thickness;
        }

        internal override void DrawShapeOnDoc(Document doc)
        {
            doc.Shapes.AddCurve(ConvertTuplePointsToSafePoints()).Line.Weight = (float)_thickness;
        }

        private object ConvertTuplePointsToSafePoints()
        {
            var temp = new Single[_points.Length, 2];
            for (int i = 0; i < _points.Length; i++)
            {
                var (x, y) = _points[i];
                temp[i, 0] = (Single)x;
                temp[i, 1] = (Single)y;
            }
            return temp;
        }
    }
}