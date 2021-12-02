using EMDD.Reporting.Line;

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
        public int ColCount { get; }

        /// <summary>
        /// Total Rows
        /// </summary>
        public int RowCount { get; }

        /// <summary>
        /// Table content
        /// </summary>
        public string[,] Content { get; }

        /// <summary>
        /// cells indexes to be merged
        /// </summary>
        public List<int[,]> CellRange2Merge { get; }

        /// <summary>
        /// Iniitialize
        /// </summary>
        /// <param name="pContent"></param>
        public LineTable(string[,] pContent, uint tabLevel) : base(tabLevel)
        {
            Content = pContent;
            RowCount = Content.GetUpperBound(0);
            ColCount = Content.GetUpperBound(1);
            CellRange2Merge = new List<int[,]>();
        }

        /// <summary>
        /// Merge Row Range
        /// </summary>
        /// <param name="prow"></param>
        /// <param name="pcol1"></param>
        /// <param name="pcol2"></param>
        public void RowMerge(int prow, int pcol1, int pcol2)
        {
            var limrow = prow.LimitWithin(0, RowCount);
            var limcol1 = pcol1.LimitWithin(0, ColCount);
            var limcol2 = pcol2.LimitWithin(0, ColCount);
            var tempRange = new[,] {
                    {limrow, limcol1},
                    {limrow, limcol2}
                };
            CellRange2Merge.Add(tempRange);
        }

        /// <summary>
        /// Merge Column Range
        /// </summary>
        /// <param name="pcol"></param>
        /// <param name="prow1"></param>
        /// <param name="prow2"></param>
        public void ColMerge(int pcol, int prow1, int prow2)
        {
            var limcol = pcol.LimitWithin(0, ColCount);
            var limrow1 = prow1.LimitWithin(0, RowCount);
            var limrow2 = prow2.LimitWithin(0, RowCount);
            var tempRange = new[,] {
                    {limrow1, limcol},
                    {limrow2, limcol}
                };
            CellRange2Merge.Add(tempRange);
        }

        internal override void WriteToString(ref StringBuilder str)
        {
            for (int i = 0; i < RowCount; i++)
            {
                str.Append(new string('\t', (int)TabIndex));
                for (int j = 0; j < ColCount-1; j++)
                {
                    str.Append('|').Append(Content[i, j]).Append(new string('\t', 2));
                }
                str.Append('|').Append(Content[i, ColCount - 1]).Append('|').AppendLine();
            }
        }
    }
}