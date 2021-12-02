using System.Text;

namespace EMDD.Reporting
{
    /// <summary>
    /// Composition of paragraphs
    /// </summary>
    public class Essay
    {
        public string Title { get; }

        /// <summary>
        /// Initialize with Title (heading)
        /// </summary>
        /// <param name="pTitle"></param>
        public Essay(string pTitle)
        {
            Title = pTitle;
            Paragraphs = new List<Paragraph>();
            NewParagraph("", 0, 0);
        }

        public List<Paragraph> Paragraphs { get; }

        /// <summary>
        /// Create new Paragraph
        /// </summary>
        /// <param name="pTitle"></param>
        public Paragraph NewParagraph(string pTitle, int tabSpace, uint tabIndex)
        {
            var _currentParagraph = new Paragraph(pTitle, tabSpace, tabIndex);
            Paragraphs.Add(_currentParagraph);
            return _currentParagraph;
        }

        public void AddParagraph(Paragraph paragraph)
        {
            Paragraphs.Add(paragraph);
        }

        public string ToCompleteString()
        {
            var str = new StringBuilder();
            str.AppendLine(Title);
            foreach (var par in Paragraphs)
            {
                par.WriteToString(ref str);
            }
            return str.ToString();
        }
    }
}