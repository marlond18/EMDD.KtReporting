using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace Reporting
{
    /// <summary>
    /// Composition of paragraphs
    /// </summary>
    public class Essay
    {
        private readonly string _title;

        /// <summary>
        /// Initialize with Title (heading)
        /// </summary>
        /// <param name="pTitle"></param>
        public Essay(string pTitle)
        {
            _title = string.IsNullOrEmpty(pTitle) ? " " : pTitle;
            _paragraphs = new List<Paragraph>();
            NewParagraph();
        }

        private readonly List<Paragraph> _paragraphs;
        private Paragraph _currentParagraph;

        /// <summary>
        /// Create new Paragraph
        /// </summary>
        /// <param name="pTitle"></param>
        public void NewParagraph(string pTitle = " ")
        {
            _currentParagraph = new Paragraph(pTitle);
            _paragraphs.Add(_currentParagraph);
        }

        /// <summary>
        /// Append a canvas to the current paragraph
        /// </summary>
        /// <param name="val"></param>
        public void Append(LineCanvas val)
        {
            _currentParagraph.AddCanvas(val);
        }

        /// <summary>
        /// Add text to the current paragraph
        /// </summary>
        /// <param name="val"></param>
        public void Append(string val)
        {
            _currentParagraph.AddText(val);
        }

        /// <summary>
        /// Add Picture to the current paragraph
        /// </summary>
        /// <param name="val"></param>
        public void Append(Bitmap val)
        {
            _currentParagraph.AddPicture(val);
        }

        /// <summary>
        /// Add table to the current paragraph using array
        /// </summary>
        /// <param name="val"></param>
        public void Append(string[,] val)
        {
            _currentParagraph.AddTable(val);
        }

        private Word.Application _oApp;

        /// <summary>
        /// Close the related application
        /// </summary>
        public void KillApp()
        {
            try
            {
                if (_oApp?.Documents.Count > 0) _oApp?.Quit();
                Marshal.FinalReleaseComObject(_oApp);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// Create the Word Document and write the pertinent items to it
        /// </summary>
        /// <param name="visible"></param>
        public void CreateWordDoc(bool visible = true)
        {
            try
            {
                _oApp = new Word.Application { Visible = visible };
                var oDoc = _oApp.Documents.Add();
                oDoc.PageSetup.TopMargin = 40;
                oDoc.PageSetup.BottomMargin = 40;
                oDoc.PageSetup.LeftMargin = 40;
                oDoc.PageSetup.RightMargin = 40;
                var oParag = oDoc.Content.Paragraphs.Add();
                (new LineText(_title)).CreateLine(oParag.Range, WdOMathJc.wdOMathJcCenter, 20, 0, 10);
                foreach (var paragraph in _paragraphs)
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
}