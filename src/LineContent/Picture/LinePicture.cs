using EMDD.Reporting.Line;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace EMDD.Reporting
{
    /// <summary>
    /// Picture Line for Paragraph
    /// </summary>
    public class LinePicture : LineContent
    {
        /// <summary>
        /// The Bitmap to be placed on word
        /// </summary>
        private readonly Bitmap _pictureContent;

        /// <summary>
        /// Initialize with picture
        /// </summary>
        /// <param name="pContent"></param>
        public LinePicture(Bitmap pContent, uint tabLevel):base(tabLevel)
        {
            _pictureContent = pContent;
        }

        /// <summary>
        /// Initialize with control
        /// </summary>
        /// <param name="pBox"></param>
        public LinePicture(Control pBox, uint tabLevel) : base(tabLevel)
        {
            var bmp = new Bitmap(pBox.Width, pBox.Height);
            pBox.DrawToBitmap(bmp, pBox.ClientRectangle);
            _pictureContent = bmp;
        }

        internal override void WriteLine(Range range, WdOMathJc justify = WdOMathJc.wdOMathJcLeft, int fontsize = 12, int leftIndent = 0, int spaceAfter = 0, int bold = 0)
        {
            if (_pictureContent == null) return;
            Clipboard.SetImage(_pictureContent);
            range.Paste();
        }
    }
}