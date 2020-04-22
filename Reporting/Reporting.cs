using KtExtensions.NetStandard;

namespace Reporting
{
    /// <summary>
    /// Symbols
    /// </summary>
    public static class EquivText
    {
        /// <summary>
        /// Gamma Symbol
        /// </summary>
        public const string Gamma = "γ";

        /// <summary>
        /// Multipier Symbol
        /// </summary>
        public const string Times = "×";
    }

    /// <summary>
    /// String builder of Math Equations
    /// </summary>
    public static class MathStringBuilder
    {
        /// <summary>
        /// Matrix
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ToMatrix(this string[,] str) => "[■(" + str.SelectRows(row => row.BuildString("&")).BuildString("@") + ")]";

        /// <summary>
        /// intergral sign
        /// </summary>
        /// <param name="exp"></param>
        /// <param name="lower"></param>
        /// <param name="upper"></param>
        /// <param name="variableOfInt"></param>
        /// <returns></returns>
        public static string Integrate(string exp, string lower, string upper, string variableOfInt) => $"∫_{lower}^{upper}▒({exp})d{variableOfInt}";

        /// <summary>
        /// summation
        /// </summary>
        /// <param name="exp"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="variable"></param>
        /// <returns></returns>
        public static string Sigma(string exp, string from, string to, string variable) => $"∑_({variable}={from})^{to}▒{exp}";
    }
}