using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace EMDD.Reporting
{
    /// <summary>
    /// Composition of paragraphs
    /// </summary>
    public class Essay
    {
        private readonly string _title;

        public float TopMargin { get; set; } = 40;
        public float BottomMargin { get; set; } = 40;
        public float LeftMargin { get; set; } = 40;
        public float RightMargin { get; set; } = 40;

        /// <summary>
        /// Initialize with Title (heading)
        /// </summary>
        /// <param name="pTitle"></param>
        public Essay(string pTitle)
        {
            _title = pTitle;
            _paragraphs = new List<Paragraph>();
            NewParagraph("", 0, 0);
        }

        private readonly List<Paragraph> _paragraphs;

        /// <summary>
        /// Create new Paragraph
        /// </summary>
        /// <param name="pTitle"></param>
        public Paragraph NewParagraph(string pTitle, int tabSpace, uint tabIndex)
        {
            var _currentParagraph = new Paragraph(pTitle, tabSpace, tabIndex);
            _paragraphs.Add(_currentParagraph);
            return _currentParagraph;
        }

        public void AddParagraph(Paragraph paragraph)
        {
            _paragraphs.Add(paragraph);
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
        public void CreateWordDoc()
        {
            try
            {
                _oApp = new Word.Application { Visible = true };
                var oDoc = _oApp.Documents.Add();
                oDoc.PageSetup.TopMargin = TopMargin;
                oDoc.PageSetup.BottomMargin = BottomMargin;
                oDoc.PageSetup.LeftMargin = LeftMargin;
                oDoc.PageSetup.RightMargin = RightMargin;
                var oParag = oDoc.Content.Paragraphs.Add();
                if (!string.IsNullOrEmpty(_title) && !string.IsNullOrWhiteSpace(_title))
                    new LineText(_title, 0).CreateLine(oParag.Range, WdOMathJc.wdOMathJcCenter, 20, 0, 10);
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