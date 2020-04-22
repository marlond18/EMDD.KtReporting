using KtExtensions;
using Reporting.Line;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace Reporting
{
    /// <summary>
    /// Compositions of Lines
    /// </summary>
    public class Paragraph
    {
        private readonly string _title;

        /// <summary>
        /// Initialize paragraph with the title
        /// </summary>
        /// <param name="pTitle"></param>
        public Paragraph(string pTitle)
        {
            _title = string.IsNullOrEmpty(pTitle) ? " " : pTitle;
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
        public void AddText(string val)
        {
            Content.Add(new LineText(val));
        }

        /// <summary>
        /// Add Picture to the paragraph using bitmap
        /// </summary>
        /// <param name="val"></param>
        public void AddPicture(Bitmap val)
        {
            Content.Add(new LinePicture(val));
        }

        /// <summary>
        /// Add picture to the paragraph using picture box of the name windows picturebox
        /// </summary>
        /// <param name="val"></param>
        public void AddPicture(PictureBox val)
        {
            Content.Add(new LinePicture(val));
        }

        /// <summary>
        /// Add table to the paragraph using array
        /// </summary>
        /// <param name="val"></param>
        public void AddTable(string[,] val)
        {
            Content.Add(new LineTable(val));
        }

        /// <summary>
        /// add the table to the paragraph using line table
        /// </summary>
        /// <param name="val"></param>
        public void AddTable(LineTable val)
        {
            Content.Add(val);
        }

        /// <summary>
        /// add shape canvas to the paragraph
        /// </summary>
        /// <param name="val"></param>
        public void AddCanvas(LineCanvas val)
        {
            Content.Add(val);
        }

        internal void WriteParagraph(Word.Paragraph oParag)
        {
            (new LineText(_title)).CreateLine(oParag.Range);
            foreach (var line in Content)
            {
                line.CreateLine(oParag.Range, WdOMathJc.wdOMathJcLeft, 10, 15);
            }
        }
    }
}