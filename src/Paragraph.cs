using EMDD.Reporting.Line;

using System.Drawing;
using System.Text;

namespace EMDD.Reporting
{
    /// <summary>
    /// Compositions of Lines
    /// </summary>
    public class Paragraph
    {
        /// <summary>
        /// Initialize paragraph with the title
        /// </summary>
        /// <param name="title"></param>
        /// <param name="defaulttab"></param>
        /// <param name="tabIndex"></param>
        public Paragraph(string title, int defaulttab, uint tabIndex)
        {
            Content = new List<LineContent>();
            Title = title;
            Defaulttab = defaulttab;
            TabIndex = tabIndex;
        }

        /// <summary>
        /// The collection of the paragraph contents
        /// </summary>
        public List<LineContent> Content { get; }
        public string Title { get; }
        public int Defaulttab { get; }
        public uint TabIndex { get; }

        /// <summary>
        /// Add text to the paragraph
        /// </summary>
        /// <param name="val"></param>
        public void AddText(string val, uint tabIndex)
        {
            Content.Add(new LineText(val, tabIndex + TabIndex));
        }

        internal void WriteToString(ref StringBuilder str)
        {
            str.Append(new string('\t', (int)TabIndex)).AppendLine(Title);
            foreach (var line in Content)
            {
                line.WriteToString(ref str);
            }
        }

        /// <summary>
        /// Add Picture to the paragraph using bitmap
        /// </summary>
        /// <param name="val"></param>
        public void AddPicture(Bitmap val, uint tabIndex)
        {
            Content.Add(new LinePicture(val, tabIndex + TabIndex));
        }

        /// <summary>
        /// Add table to the paragraph using array
        /// </summary>
        /// <param name="val"></param>
        public void AddTable(string[,] val, uint tabIndex)
        {
            Content.Add(new LineTable(val, tabIndex + TabIndex));
        }

        ///// <summary>
        ///// add the table to the paragraph using line table
        ///// </summary>
        ///// <param name="val"></param>
        //public void AddTable(LineTable val)
        //{
        //    Content.Add(val);
        //}

        ///// <summary>
        ///// add shape canvas to the paragraph
        ///// </summary>
        ///// <param name="val"></param>
        //public void AddCanvas(LineCanvas val)
        //{
        //    Content.Add(val);
        //}

        ///// <summary>
        ///// Add Generic LineContemt
        ///// </summary>
        ///// <param name="line"></param>
        //public void AddLineContent<T>(T line) where T : LineContent
        //{
        //    Content.Add(line);
        //}
    }
}