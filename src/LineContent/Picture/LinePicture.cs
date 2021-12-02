using EMDD.Reporting.Line;

using System.Drawing;

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
        public Bitmap PictureContent { get; }

        /// <summary>
        /// Initialize with picture
        /// </summary>
        /// <param name="pContent"></param>
        public LinePicture(Bitmap pContent, uint tabLevel):base(tabLevel)
        {
            PictureContent = pContent;
        }

        internal override void WriteToString(ref StringBuilder str)
        {
            str.Append(new string('\t',(int)TabIndex)).AppendLine("<Picture Not converted To basic String>");
        }
    }
}