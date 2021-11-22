using EMDD.Reporting.Line;
using Microsoft.Office.Interop.Word;
using KtExtensions;

namespace EMDD.Reporting
{
    /// <summary>
    /// Line Content Text
    /// </summary>
    public class LineText : LineContent
    {
        /// <summary>
        /// Text Content
        /// </summary>
        private readonly string _textContent;

        /// <summary>
        /// initialize
        /// </summary>
        /// <param name="tContent"></param>
        public LineText(string tContent, uint tabLevel) : base(tabLevel)
        {
            _textContent = tContent.IsNull() || tContent.IsEmpty() ? " " : tContent;
        }

        /// <summary>
        /// LineText to Text
        /// </summary>
        /// <param name="text"></param>
        public static implicit operator string(LineText text) => text._textContent;

        internal override void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
            if ((_textContent == null) || (_textContent?.Length == 0)) return;
            range.Text = _textContent;
            var mathRange = range.OMaths.Add(range);
            var currentMath = mathRange.OMaths[1];
            currentMath.Range.Font.Bold = bold;
            currentMath.Range.Font.Size = fontsize;
            currentMath.Justification = justify;
            currentMath.BuildUp();
            currentMath.Range.Paragraphs.LeftIndent = leftIndent;
            currentMath.Range.Paragraphs.SpaceAfter = spaceAfter;
        }
    }
}