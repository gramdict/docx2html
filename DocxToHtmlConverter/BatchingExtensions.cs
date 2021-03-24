using System.Collections.Generic;

namespace DocxToHtmlConverter
{
    static class BatchingExtensions
    {
        public static IEnumerable<IEnumerable<string>> BatchByLen(this IEnumerable<string> strings, int maxLen)
        {
            var bucket = new List<string>();
            var len = 0;
            
            foreach (var s in strings)
            {
                if (len + s.Length > maxLen)
                {
                    yield return bucket;
                    bucket.Clear();
                    len = 0;
                }

                bucket.Add(s);
                len += s.Length;
            }

            if (bucket.Count > 0)
                yield return bucket;
        }
    }
}