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
            Vertex = (x, y);
            Width = w;
            Height = h;
        }

        public object CreatePolyLine()
        {
            var arr = new Single[4, 2];
            arr[0, 0] = (float)(Vertex.x);
            arr[0, 1] = (float)(Vertex.y);
            arr[1, 0] = (float)(Vertex.x + (Width / 2));
            arr[1, 1] = (float)(Vertex.y + Height);
            arr[2, 0] = (float)(Vertex.x - (Width / 2));
            arr[2, 1] = (float)(Vertex.y + Height);
            arr[3, 0] = (float)(Vertex.x);
            arr[3, 1] = (float)(Vertex.y);
            return arr;
        }

        public  double Width { get; }
        public  double Height { get; }
        public (double x, double y) Vertex { get; }
    }
}