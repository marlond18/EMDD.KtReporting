using EMDD.Reporting.Line;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    /// <summary>
    /// Compositions of Lines
    /// </summary>
    public class Paragraph
    {
        private readonly string _title;

        private readonly int _defaulttab;

        private readonly uint _tabIndex;

        /// <summary>
        /// Initialize paragraph with the title
        /// </summary>
        /// <param name="pTitle"></param>
        public Paragraph(string pTitle, int defaulttab, uint tabIndex)
        {
            _title = pTitle;
            _defaulttab = defaulttab;
            _tabIndex = tabIndex;
            Content = new List<LineContent>();
        }

        /// <summary>
        /// The collection of the paragraph contents
        /// </summary>
        internal List<LineContent> Content { get; }

        /// <summary>
        /// Add text to the paragraph
        /// </summary>
        /// <param name="val"></param>
        public void AddText(string val, uint tabIndex)
        {
            Content.Add(new LineText(val, tabIndex + _tabIndex));
        }

        /// <summary>
        /// Add Picture to the paragraph using bitmap
        /// </summary>
        /// <param name="val"></param>
        public void AddPicture(Bitmap val, uint tabIndex)
        {
            Content.Add(new LinePicture(val, tabIndex + _tabIndex));
        }

        /// <summary>
        /// Add picture to the paragraph using picture box of the name windows picturebox
        /// </summary>
        /// <param name="val"></param>
        public void AddPicture(PictureBox val, uint tabIndex)
        {
            Content.Add(new LinePicture(val, tabIndex + _tabIndex));
        }

        /// <summary>
        /// Add table to the paragraph using array
        /// </summary>
        /// <param name="val"></param>
        public void AddTable(string[,] val, uint tabIndex)
        {
            Content.Add(new LineTable(val, tabIndex + _tabIndex));
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

        internal void WriteParagraph(Word.Paragraph oParag)
        {
            if (!string.IsNullOrEmpty(_title) && !string.IsNullOrWhiteSpace(_title))
                new LineText(_title, 0).CreateLine(oParag.Range);
            foreach (var line in Content)
            {
                line.CreateLine(oParag.Range, WdOMathJc.wdOMathJcLeft, 10, _defaulttab);
            }
        }
    }
}