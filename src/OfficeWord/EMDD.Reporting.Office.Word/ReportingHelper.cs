using EMDD.Reporting.Line;

using Microsoft.Office.Interop.Word;

using System.Runtime.InteropServices;

using MWord = Microsoft.Office.Interop.Word;

namespace EMDD.Reporting.Office.Word
{
    public class WordEssay
    {
        public WordEssay(Essay essay)
        {
            Essay = essay;
        }

        public float TopMargin { get; set; } = 40;
        public float BottomMargin { get; set; } = 40;
        public float LeftMargin { get; set; } = 40;
        public float RightMargin { get; set; } = 40;

        private MWord.Application? _oApp;

        public Essay Essay { get; }

        /// <summary>
        /// Close the related application
        /// </summary>
        public void KillApp()
        {
            try
            {
                if (_oApp?.Documents.Count > 0) _oApp?.Quit();
                if(_oApp is not null) Marshal.FinalReleaseComObject(_oApp);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// Create the Word Document and write the pertinent items to it
        /// </summary>
        public void CreateWordDoc()
        {
            try
            {
                _oApp = new MWord.Application { Visible = true };
                var oDoc = _oApp.Documents.Add();
                oDoc.PageSetup.TopMargin = TopMargin;
                oDoc.PageSetup.BottomMargin = BottomMargin;
                oDoc.PageSetup.LeftMargin = LeftMargin;
                oDoc.PageSetup.RightMargin = RightMargin;
                var oParag = oDoc.Content.Paragraphs.Add();
                if (!string.IsNullOrEmpty(Essay.Title) && !string.IsNullOrWhiteSpace(Essay.Title))
                    new LineText(Essay.Title, 0).CreateLine(oParag.Range, WdOMathJc.wdOMathJcCenter, 20, 0, 10);
                foreach (var paragraph in Essay.Paragraphs)
                {
                    paragraph.WriteParagraph(oParag);
                }
                _oApp.Visible = true;
            }
            catch (Exception ex)
            {
                KillApp();
                MessageBox.Show(ex.Message);
            }
        }
    }

    public static class ReportingHelper
    {
        public static Bitmap ToBitmap(this Control pBox)
        {
            var bmp = new Bitmap(pBox.Width, pBox.Height);
            pBox.DrawToBitmap(bmp, pBox.ClientRectangle);
            return bmp;
        }

        /// <summary>
        /// Initialize with control
        /// </summary>
        /// <param name="pBox"></param>
        public static LinePicture LinePicFromPicBox(Control pBox, uint tabLevel)
        {
            return new LinePicture(pBox.ToBitmap(), tabLevel);
        }

        public static void WriteParagraph(this Paragraph parag, MWord.Paragraph oParag)
        {
            if (!string.IsNullOrEmpty(parag.Title) && !string.IsNullOrWhiteSpace(parag.Title))
                new LineText(parag.Title, 0).CreateLine(oParag.Range);
            foreach (var line in parag.Content)
            {
                line.CreateLine(oParag.Range, WdOMathJc.wdOMathJcLeft, 10, parag.Defaulttab);
            }
        }

        private static void WriteLine(this LineCanvas lc, MWord.Range range)
        {
            var canvas = range.Document.Shapes.AddCanvas((float)lc.Location.X, (float)lc.Location.Y, (float)lc.Size.Width, (float)lc.Size.Height, range);
            canvas.WrapFormat.Type = WdWrapType.wdWrapInline;
            foreach (var shape in lc.Shapes)
            {
                switch (shape)
                {
                    case CurveCanvasShape ccs:
                        ccs.DrawShapeOnCanvas(canvas.CanvasItems);
                        break;
                    case LineCanvasShape lcs:
                        lcs.DrawShapeOnCanvas(canvas.CanvasItems);
                        break;
                    case TriangleCanvasShape tcs:
                        tcs.DrawShapeOnCanvas(canvas.CanvasItems);
                        break;
                }
            }
            canvas.CanvasItems.SelectAll();
            range.Application.Selection.Cut();
            canvas.Delete();
            range.PasteSpecial(DataType: WdPasteDataType.wdPasteEnhancedMetafile);
        }

        private static void DrawShapeOnCanvas(this TriangleCanvasShape tcs, MWord.CanvasShapes canvasItems)
        {
            var shp = canvasItems.AddPolyline(tcs.CreatePolyLine());
            shp.Fill.BackColor.SchemeColor = 2;
        }

        private static void DrawShapeOnDoc(this TriangleCanvasShape tcs, Document doc)
        {
            var shp = doc.Shapes.AddPolyline(tcs.CreatePolyLine());
            shp.Fill.BackColor.SchemeColor = 2;
        }

        private static void DrawShapeOnDoc(this LineCanvasShape lcs, Document doc)
        {
            var line = doc.Shapes.AddLine((float)lcs.Start.x, (float)lcs.Start.y, (float)lcs.End.x, (float)lcs.End.y);
            line.Line.Weight = (float)lcs.Thickness;
        }

        private static void DrawShapeOnCanvas(this LineCanvasShape lcs, MWord.CanvasShapes canvasItems)
        {
            var line = canvasItems.AddLine((float)lcs.Start.x, (float)lcs.Start.y, (float)lcs.End.x, (float)lcs.End.y);
            line.Line.Weight = (float)lcs.Thickness;
        }

        private static void DrawShapeOnCanvas(this CurveCanvasShape ccs, MWord.CanvasShapes canvasItems)
        {
            canvasItems.AddCurve(ccs.ConvertTuplePointsToSafePoints()).Line.Weight = (float)ccs.Thickness;
        }

        private static void DrawShapeOnDoc(this CurveCanvasShape ccs, Document doc)
        {
            doc.Shapes.AddCurve(ccs.ConvertTuplePointsToSafePoints()).Line.Weight = (float)ccs.Thickness;
        }

        private static void WriteLine(this LineText lt, MWord.Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
            if ((lt.TextContent == null) || (lt.TextContent?.Length == 0)) return;
            range.Text = lt.TextContent;
            var mathRange = range.OMaths.Add(range);
            var currentMath = mathRange.OMaths[1];
            currentMath.Range.Font.Bold = bold;
            currentMath.Range.Font.Size = fontsize;
            currentMath.Justification = justify;
            currentMath.BuildUp();
            currentMath.Range.Paragraphs.LeftIndent = leftIndent;
            currentMath.Range.Paragraphs.SpaceAfter = spaceAfter;
            range.InsertParagraphAfter();
        }

        private static  void WriteLine(this LineTable lta, MWord.Range range,int fontsize = 12)
        {
            var oTable = range.Tables.Add(range.Bookmarks["\\endofdoc"].Range, lta.RowCount + 1, lta.ColCount + 1);
            for (var r = 0; r < lta.RowCount + 1; r++)
            {
                for (var c = 0; c < lta.ColCount + 1; c++)
                {
                    var cellRange = oTable.Cell(r + 1, c + 1).Range;
                    cellRange.Text = lta.Content[r, c];
                    cellRange.OMaths.Add(cellRange);
                    cellRange.OMaths[1].BuildUp();
                    cellRange.OMaths[1].Range.Font.Size = fontsize;
                }
                foreach (var merge in lta.CellRange2Merge)
                {
                    var cell1 = oTable.Cell(merge[0, 0] + 1, merge[0, 1] + 1);
                    var cell2 = oTable.Cell(merge[1, 0] + 1, merge[1, 1] + 1);
                    cell1.Merge(cell2);
                }
                oTable.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleOutset;
                oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            }
        }

        private static void WriteLine(this LinePicture lp, MWord.Range range)
        {
            if (lp.PictureContent == null) return;
            Clipboard.SetImage(lp.PictureContent);
            range.Paste();
        }

        internal static void CreateLine<T>(this T lc, MWord.Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0) where T: LineContent
        {
            switch (lc)
            {
                case LineText lt:
                    lt.WriteLine(range, justify, fontsize, leftIndent * (int)lc.TabIndex, spaceAfter, bold);
                    break;
                case LineTable lta:
                    lta.WriteLine(range, fontsize);
                    break;
                case LinePicture lp:
                    lp.WriteLine(range);
                    break;
                case LineCanvas ls:
                    ls.WriteLine(range);
                    break;
                case CurveCanvasShape ccs:
                    ccs.DrawShapeOnDoc(range.Document);
                    break;
                case LineCanvasShape lcs:
                    lcs.DrawShapeOnDoc(range.Document);
                    break;
                case TriangleCanvasShape tcs:
                    tcs.DrawShapeOnDoc(range.Document);
                    break;
            }
            range.InsertParagraphAfter();
        }
    }
}