namespace EMDD.Reporting
{
    /// <summary>
    /// Collection of Math Display Shortcuts
    /// </summary>
    public static class MathDisplay
    {
        /// <summary>
        /// Summation statement
        /// </summary>
        /// <param name="lowerLimit"> starting value, ex. i=1 </param>
        /// <param name="upperLimit">total iterations,ex. N=10</param>
        /// <param name="expression">the expression inside the summation operation</param>
        /// <returns></returns>
        public static string Sigma(string lowerLimit, string upperLimit, string expression)
        {
            return $"∑_({lowerLimit})^({upperLimit})▒({expression}) ";
        }
    }
}