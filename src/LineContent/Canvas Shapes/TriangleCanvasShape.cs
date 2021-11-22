using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System;

namespace EMDD.Reporting
{
    /// <summary>
    /// Triangles for canvas
    /// </summary>
    public class TriangleCanvasShape : CanvasShapes
    {
        /// <summary>
        /// Initialize
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="w"></param>
        /// <param name="h"></param>
        public TriangleCanvasShape(double x, double y, double w, double h, uint tabLevel):base(tabLevel)
        {
            _vertex = (x, y);
            _width = w;
            _height = h;
        }

        internal override void DrawShapeOnCanvas(Word.CanvasShapes canvasItems)
        {
            var shp = canvasItems.AddPolyline(CreatePolyLine());
            shp.Fill.BackColor.SchemeColor = 2;
        }

        private object CreatePolyLine()
        {
            var arr = new Single[4, 2];
            arr[0, 0] = (float)(_vertex.x);
            arr[0, 1] = (float)(_vertex.y);
            arr[1, 0] = (float)(_vertex.x + (_width / 2));
            arr[1, 1] = (float)(_vertex.y + _height);
            arr[2, 0] = (float)(_vertex.x - (_width / 2));
            arr[2, 1] = (float)(_vertex.y + _height);
            arr[3, 0] = (float)(_vertex.x);
            arr[3, 1] = (float)(_vertex.y);
            return arr;
        }

        internal override void DrawShapeOnDoc(Document doc)
        {
            var shp = doc.Shapes.AddPolyline(CreatePolyLine());
            shp.Fill.BackColor.SchemeColor = 2;
        }

        private readonly double _width;
        private readonly double _height;
        private (double x, double y) _vertex;
    }
}