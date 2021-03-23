using EMDD.Reporting.Line;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    /// <summary>
    /// base shape for canvas
    /// </summary>
    public abstract class CanvasShapes : LineContent
    {
        internal override void WriteLine(Word.Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
            DrawShapeOnDoc(range.Document);
        }

        internal abstract void DrawShapeOnCanvas(Word.CanvasShapes canvasItems);
        internal abstract void DrawShapeOnDoc(Document doc);
    }
}