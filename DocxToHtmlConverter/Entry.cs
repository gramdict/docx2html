using System.Collections.Generic;

namespace DocxToHtmlConverter
{
    public class Entry
    {
        public string Lemma;
        public IEnumerable<Definition> Definitions;
    }
}