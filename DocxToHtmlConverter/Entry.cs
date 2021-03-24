using System.Collections.Generic;

namespace DocxToHtmlConverter
{
    public class Entry
    {
        public string Numbers;
        public string Lemma;
        public string Parens;
        public IEnumerable<Definition> Definitions;
    }
}