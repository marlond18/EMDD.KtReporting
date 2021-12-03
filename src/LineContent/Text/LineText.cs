using EMDD.Reporting.Line;

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
        public string TextContent { get; }

        /// <summary>
        /// initialize
        /// </summary>
        /// <param name="tContent"></param>
        public LineText(string tContent, uint tabLevel) : base(tabLevel)
        {
            TextContent = tContent.IsNull() || tContent.IsEmpty() ? " " : tContent;
        }

        /// <summary>
        /// LineText to Text
        /// </summary>
        /// <param name="text"></param>
        public static implicit operator string(LineText text) => text.TextContent;

        public override void WriteToString(ref StringBuilder str)
        {
            str.Append(new string('\t', (int)TabIndex)).AppendLine(TextContent);
        }
    }
}