using EMDD.Reporting.Line;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using KtExtensions;

namespace EMDD.Reporting
{
    /// <summary>
    /// Create Table
    /// </summary>
    public class LineTable : LineContent
    {
        /// <summary>
        /// Total Columns
        /// </summary>
        private readonly int _colCount;

        /// <summary>
        /// Total Rows
        /// </summary>
        private readonly int _rowCount;

        /// <summary>
        /// Table content
        /// </summary>
        private readonly string[,] _content;

        /// <summary>
        /// cells indexes to be merged
        /// </summary>
        private readonly List<int[,]> _cellRange2Merge;

        /// <summary>
        /// Iniitialize
        /// </summary>
        /// <param name="pContent"></param>
        public LineTable(string[,] pContent)
        {
            _content = pContent;
            _rowCount = _content.GetUpperBound(0);
            _colCount = _content.GetUpperBound(1);
            _cellRange2Merge = new List<int[,]>();
        }

        /// <summary>
        /// Merge Row Range
        /// </summary>
        /// <param name="prow"></param>
        /// <param name="pcol1"></param>
        /// <param name="pcol2"></param>
        public void RowMerge(int prow, int pcol1, int pcol2)
        {
            var limrow = prow.LimitWithin(0, _rowCount);
            var limcol1 = pcol1.LimitWithin(0, _colCount);
            var limcol2 = pcol2.LimitWithin(0, _colCount);
            var tempRange = new[,] {
                    {limrow, limcol1},
                    {limrow, limcol2}
                };

            _cellRange2Merge.Add(tempRange);
        }

        /// <summary>
        /// Merge Column Range
        /// </summary>
        /// <param name="pcol"></param>
        /// <param name="prow1"></param>
        /// <param name="prow2"></param>
        public void ColMerge(int pcol, int prow1, int prow2)
        {
            var limcol = pcol.LimitWithin(0, _colCount);
            var limrow1 = prow1.LimitWithin(0, _rowCount);
            var limrow2 = prow2.LimitWithin(0, _rowCount);
            var tempRange = new[,] {
                    {limrow1, limcol},
                    {limrow2, limcol}
                };
            _cellRange2Merge.Add(tempRange);
        }

        internal override void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
            var oTable = range.Tables.Add(range.Bookmarks["\\endofdoc"].Range, _rowCount + 1, _colCount + 1);
            for (var r = 0; r < _rowCount + 1; r++)
            {
                for (var c = 0; c < _colCount + 1; c++)
                {
                    var cellRange = oTable.Cell(r + 1, c + 1).Range;
                    cellRange.Text = _content[r, c];
                    cellRange.OMaths.Add(cellRange);
                    cellRange.OMaths[1].BuildUp();
                    cellRange.OMaths[1].Range.Font.Size = fontsize;
                }
                foreach (var merge in _cellRange2Merge)
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
    }
}