using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocxToHtmlConverter
{
    public static class LemmaExtensions
    {
        public static string ConvertStressMarksToNumbers(this string lemma)
        {
            MatchCollection matches = regex.Matches(lemma);

            string strippedLemma = lemma
                .Replace("\u0300", "")
                .Replace("\u0301", "")
                .Replace('ё', 'е')
                .Replace('Ё', 'Е');

            string result = strippedLemma
                            + " "
                            + string.Join("+", matches.Select(m => m.Value).Select(GetMarks));
            return result;
        }

        private static readonly Regex regex = new(@"\b[\w’]+\b");

        private static string GetMarks(string lemma)
        {
            var primary = new List<int>();
            var secondary = new List<int>();
            var yo = new List<int>();

            for (int i = 0; i < lemma.Length; ++i)
            {
                char c = lemma[i];

                if (c == 'ё' || c == 'Ё') yo.Add(i - primary.Count - secondary.Count + 1);
                if (c == '\u0301') primary.Add(i - primary.Count - secondary.Count);
                if (c == '\u0300') secondary.Add(i - primary.Count - secondary.Count);
            }

            int strippedLemmaLength = lemma.Length - primary.Count - secondary.Count;

            string marks = primary.Count switch
                           {
                               0 => "0",
                               _ => string.Join("//", primary.OrderByDescending(_ => _).Select(p => (strippedLemmaLength - p + 1).ToString()))
                           }
                           + string.Join("", yo.Select(i => "," + (strippedLemmaLength - i + 1)))
                           + string.Join("", secondary.Select(i => "." + i));
            return marks;
        }
    }
}