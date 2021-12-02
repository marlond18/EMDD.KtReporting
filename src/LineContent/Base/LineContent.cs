using System.Text;

namespace EMDD.Reporting
{
    namespace Line
    {
        /// <summary>
        /// Abstract Class for Line of Texts
        /// </summary>
        public abstract class LineContent
        {
            protected LineContent(uint tabIndex)
            {
                TabIndex = tabIndex;
            }

            public uint TabIndex { get; }

            internal abstract void WriteToString(ref StringBuilder str);
        }
    }
}