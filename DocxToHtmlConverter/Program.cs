using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace DocxToHtmlConverter
{
    public static partial class Program
    {
        static void Main(string [] args)
        {
            var stopwatch = Stopwatch.StartNew();

            List<Entry> names = ConvertNames(File.ReadLines(@"../../../names.txt")).ToList();
            List<Entry> common = ConvertCommonPart(GetCleanHtml(@"../../../all.html")).ToList();

            foreach (Config config in args.SelectMany(ReadConfig))
            {
                var entries = config.Set switch
                {
                    "names" => names,
                    "common" => common,
                    _ => throw new Exception()
                };

                Func<Entry, string> format = config.Format switch
                {
                    "html" => ToOutputFormat,
                    "txt" => ToOutputFormat2,
                    _ => throw new Exception()
                };

                IEnumerable<string> lines = entries.Select(format);

                Console.WriteLine("Saving " + config.Path);
                
                switch (config.Layout)
                {
                    case "one-file":
                        File.WriteAllLines(config.Path, lines);
                        break;
                    case "four-files":
                        SaveAsFourFiles(lines, config.Path);
                        break;
                    case "AZ":
                        SaveAsAZ(entries, config.Path, format);
                        break;
                    default:
                        throw new Exception();
                }
            }

            Console.WriteLine($"{stopwatch.Elapsed.TotalMilliseconds:N0} ms");
        }

        private static void SaveAsFourFiles(IEnumerable<string> lines, string path)
        {
            var batches = lines
                .BatchByLen(565000)
                .Select((b, i) => new {Num = i, Entries = b});
            
            foreach (var batch in batches)
            {
                File.WriteAllLines(string.Format(path, batch.Num + 1), batch.Entries);
            }
        }

        static IEnumerable<Config> ReadConfig(string path) =>
            from csvLine in File.ReadLines(path)
            let line = csvLine.Split(',')
            select new Config(line[0], line[1], line[2], line[3]);

        private static void SaveAsAZ(IEnumerable<Entry> entries, string directory, Func<Entry, string> toOutputFormat)
        {
            foreach (var group in entries.GroupBy(GetGroupName))
            {
                string path = Path.Combine(directory, group.Key + ".txt");
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllLines(path, group.Select(toOutputFormat));
            }
        }

        private static string GetGroupName(Entry e) =>
            Path.Combine(
                e.Definitions.Any(d => d.Symbol.Trim(',').EndsWith("св"))
                    ? "Глаголы"
                    : "Нарицательные",
                GetFilename(e.Lemma).ToString());

        private static char GetFilename(string lemma)
        {
            char c = char.ToUpper(GetLastLetter(lemma));
            return c == 'Ё' ? 'Е' : c;
        }

        private static char GetLastLetter(string lemma) => 
            RemoveParens(lemma)
                .TrimEnd(':')
                .TrimEnd(',')
                .TrimEnd('\u0301')
                .TrimEnd('\u0300')
                .Last();

        private static string RemoveParens(string lemma)
        {
            int i = lemma.LastIndexOf(')');
            return i == -1
                ? lemma
                : (lemma.Substring(0, lemma.IndexOf('(')) + lemma.Substring(i + 1)).Trim();
        }

        private static IEnumerable<Entry> ConvertNames(IEnumerable<string> lines) =>
            lines
                .Select(NormaliseWhitespace)
                .Except(new[]
                {
                    "(Вели́кий) Но́вгород (п +) м 1a",
                })
                .Select(line => line
                    .Replace(" см. ", " <i>см.</i> ")
                    .Replace("мн. от ", "мн. <i>от</i> ")
                )
                .Select(ParseName);

        private static IEnumerable<Entry> ConvertCommonPart(string html)
        {
            string correctHtml = CorrectHtml(html);

            File.WriteAllText("all.txt", correctHtml);

            string[] lines = correctHtml.Split(new[] {Environment.NewLine}, StringSplitOptions.None);

            IEnumerable<Entry> entries = lines
                .Where(line => !line.Contains("TODO:") && !string.IsNullOrWhiteSpace(line))
                .Where(line => !line.StartsWith("<b>споко́н</b>"))
                .Where(line => !line.StartsWith("<b>угль</b>"))
                .Where(line => !line.StartsWith("<b>жураве́ль</b>"))
                .Where(line => !line.StartsWith("<b>огнь</b>"))
                .Where(line => !line.StartsWith("<b>ви́хорь</b>"))
                .Select(Parse);
            return entries;
        }

        private static string NormaliseWhitespace(string line) => 
            Regex.Replace(line.Trim(), "\\s+", " ");

        public static string RemoveStressMarksOverNonVowels(string text)
        {
            return Regex.Replace(text, "([^аоуэыяёюеи])\u0301", "$1", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        }

        private static string GetCleanHtml(string filePath)
        {
            const string cacheFile = "clean-html-cache.txt";
            var fileInfo = new FileInfo(cacheFile);
            if (!fileInfo.Exists || fileInfo.LastWriteTimeUtc < new FileInfo(filePath).LastWriteTimeUtc)
            {
                var doc = new HtmlDocument();
                doc.Load(filePath, Encoding.UTF8);
                Console.WriteLine($"Writing {cacheFile} so subsequent runs are faster.");
                File.WriteAllLines(cacheFile, CleanHtml(doc));
            }

            return File.ReadAllText(cacheFile);
        }

        private static string ToOutputFormat(Entry e) =>
            $"{e.Lemma}|" + string.Join("<br/>", e.Definitions.Select(
                def => def.Symbol + (string.IsNullOrEmpty(def.Grammar) ? "" : $"|{def.Grammar}")));
        
        private static string ToOutputFormat2(Entry e) =>
            (string.IsNullOrEmpty(e.Numbers) ? "" : e.Numbers + "/") +
            $"{e.Lemma} " 
            + ((string.IsNullOrEmpty(e.Parens) ? "" : $"({e.Parens}) ")
            + string.Join("; ", e.Definitions.Select(
                def => def.Symbol + (string.IsNullOrEmpty(def.Grammar) ? "" : $" {def.Grammar}"))))
                .Replace(";;", ";")
                .Replace("<i>", "_")
                .Replace("</i>", "_")
                .Replace("<b>", "__")
                .Replace("</b>", "__")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .RegexReplace("<sup>(.+)</sup>", "$1/");

        static string RegexReplace(this string input, string pattern, string replacement)
            => Regex.Replace(input, pattern, replacement);
        
        static Entry Parse(string line, int number)
        {
            Match match = Regex.Match(line, @"^\s*(?<spade>♠*)\s*<b>(<sup>(?<num>[ -0123456789]+)</sup>)?(?<lemma>[- а-яА-ЯёЁ\u0300\u0301]+:?)</b>\s*(?<rest>.+)$");
            if (!match.Success) throw new Exception($"No match: ({number+1}) {line}");
            string parensSymbolGrammar = match.Groups["rest"].Value;
            (string symbolGrammar, string parens) = SplitParens(parensSymbolGrammar);
            IReadOnlyList<Definition> definitions;
            if (match.Groups["lemma"].Value.EndsWith(":"))
            {
                definitions = new [] {new Definition {Symbol = "", Grammar = symbolGrammar}};
            }
            else
            {
                definitions = NormaliseWhitespace(symbolGrammar)
                    .Split("<br/>")
                    .Select(def => ParseDefinition(line, def))
                    .ToList();
            }
            return new Entry
            {
                Numbers = match.Groups["num"].Value.Trim(),
                Lemma = match.Groups["lemma"].Value.Trim(),
                Parens = parens,
                Definitions = definitions
            };
        }

        private static Definition ParseDefinition(string line, string symbolGrammar)
        {
            Match match = Regex.Match(symbolGrammar, $@"^♠?\s*(?<symbol>{FullestSymbolsRegex}(,|:|;)?)(\s+(?<grammar>.*))?$");
            if (!match.Success) throw new Exception("Regex failed: " + line);
            var def = new Definition
            {
                Symbol = match.Groups["symbol"].Value.TrimEnd(),
                Grammar = match.Groups["grammar"].Value
            };
            return def;
        }

        private static readonly string FullestSymbolsRegex = GetFullestSymbolsRegex();
        
        static Entry ParseName(string line, int number)
        {
            string fullestSymbolsRegex = FullestSymbolsRegex;
            Match match = Regex.Match(line, $@"^(?<lemma>[-’ а-яА-ЯёЁ\u0300\u0301\(\)]+:?)\s+(?<symbol>{fullestSymbolsRegex}(,|:|;)?)(\s+(?<grammar>.+))?$");
            if (!match.Success) throw new Exception($"No match: ({number+1}) {line}");
            return new Entry
            {
                Lemma = match.Groups["lemma"].Value,
                Definitions = new []{new Definition
                {
                    Symbol = match.Groups["symbol"].Value,
                    Grammar = match.Groups["grammar"].Value
                }}
            };
        }
        
        static (string symbolGrammar, string parens) SplitParens(string parensSymbolGrammar)
        {
            string symbolGrammar;
            string parens;
            if (parensSymbolGrammar.StartsWith("("))
            {
                int closingParenIndex = parensSymbolGrammar.IndexOf(')');
                parens = parensSymbolGrammar.Substring(1, closingParenIndex - 1);
                symbolGrammar = parensSymbolGrammar.Substring(closingParenIndex + 1);
            }
            else
            {
                symbolGrammar = parensSymbolGrammar;
                parens = "";
            }

            switch (parens)
            {
                case "":

                case "<i>в знач. «иди»</i>":

                case "<i>без удар.</i>":
                case "<i>часто без удар.</i>":
                case "<i>часто без удар.</i>: еще":
                case "<i>нормально без удар.</i>":
                    break;
                default:
                    throw new Exception(parens);
            }

            return (symbolGrammar, parens);
        }

        static string GetFullestSymbolsRegex()
        {
            string formsRegex = RegexEscape("_повел. от_,_наст. 3 ед. от_".Split(','));
            // TODO нп, безл., многокр.

            string symbolsRegex = RegexEscape("ф.,мо⁺,жо⁺,м,мо,ж,жо,с,со,жо,мо-жо,мн.,мн. одуш.,мн. неод.,мн. _от_,п,мс,мс-п,числ.,числ.-п,св,нсв,св-нсв,св/нсв,н,част.,част.(_усилительная_),союз,предл.,предик.,вводн.,межд.,сравн.,§1,§2,предикативное мс,_см._"
                .Split(',')
                .OrderByDescending(s => s.Length).ToArray());

            string slashSymbolsRegex = $"(//?({symbolsRegex}))?";
            string fullSymbolsRegex = $"({symbolsRegex}){slashSymbolsRegex}{slashSymbolsRegex}";

            string genPluralRegex = RegexEscape(
                "_форма Р. мн.; других форм нет_,_форма Р. мн. от сущ. (не числ.) со знач._ «сотня»".Split(','));

            string fullestSymbolsRegex = $"({genPluralRegex}|(({formsRegex}\\s+)?{fullSymbolsRegex}))";
            return fullestSymbolsRegex;
        }

        static string RegexEscape(string[] symbols)
        {
            return "(" + string.Join("|", symbols.Select(Regex.Escape).Select(ToHtml)) + ")";
        }
        
        static string ToHtml(string s)
        {
            int i=0;
            return Regex.Replace(s, "_", _ => ++i % 2 == 1 ? "<i>" : "</i>");
        }

        private static IEnumerable<string> CleanHtml(HtmlDocument doc) =>
            doc.DocumentNode.SelectNodes("html/body/div/p")
                .Skip(8)
                .Select(ToText)
                .Where(line => line.Length > 8); // Skip section headings

        public static string ToText(HtmlNode node)
        {
            Clean(node);

            string html = node.InnerHtml;

            html = Regex.Replace(html, @"\s+", " ")
                .Replace("<i> </i>", " ");

            string text = html
                    .Replace("<i>­</i>", "")
                    .Replace("<b>­</b>", "")
                    .Replace("</b><b>", "")
                    .Replace("</i><i>", "")
                ;
            return text;
        }

        static string MapDingbats(string dingbats, HtmlNode paragraph)
        {
            return Map(dingbats, paragraph, "tÀÁGÈÂÆÃÇÄÅ", "♠①②✧⑨③⑦④⑧⑤⑥");
        }

        static string MapDingbatsDop(string dingbats, HtmlNode paragraph)
        {
            return Map(dingbats, paragraph, "р", "◑");
        }

        static string MapLucidaIcons(string dingbats, HtmlNode paragraph)
        {
            return Map(dingbats, paragraph, "", "⌧");
        }

        static string MapAntiqua(string dingbats, HtmlNode paragraph)
        {
            return Map(dingbats, paragraph, "f", "f");
        }

        static string MapLucida(string dingbats, HtmlNode paragraph)
        {
            return Map(dingbats, paragraph, " W%1", " △✕*");
        }

        static string Map(string dingbats, HtmlNode paragraph, string @from, string to)
        {
            var sb = new StringBuilder();

            foreach (var dingbat in dingbats)
            {
                int i = @from.IndexOf(dingbat);
                if (i == -1)
                    throw new Exception($"Unexpected Dingbat: {dingbat} in {paragraph.InnerText}");
                sb.Append(to[i]);
            }

            return sb.ToString();
        }

        static void Clean(HtmlNode paragraph)
        {
            var child = paragraph.FirstChild;

            while (child != null)
            {
                HtmlAttribute style = child.Attributes["style"];
                if (style != null)
                {
                    string[] styles = style.Value.Replace("\n", "").Replace("\r", "").Split(';');
                    if (styles.Contains("top:-2.0pt")) child.Name = "sup";
                    style.Value = string.Join(";", styles.Where(s => !RemoveStyle(s)));
                    if (string.IsNullOrWhiteSpace(style.Value))
                        style.Remove();
                }
                child.Attributes["lang"]?.Remove();
                child.Attributes["dir"]?.Remove();
                child.Attributes["name"]?.Remove();
                child.ChildNodes.OfType<HtmlTextNode>().ToList()
                    .ForEach(c => c.Text = c.Text.Replace("&nbsp;", " "));

                switch (child.Name)
                {
                    case "span":
                    case "st1:metricconverter":
                    case "o:p":
                    case "u":
                    case "a":
                    case "sup":
                        HandleAcc(child, "Tim_acc", "\u0301");
                        HandleAcc(child, "Tim_pob", "\u0300");
                        if (child.Attributes["style"]?.Value == "font-family:\"ZapfDingbats BT\"")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().FirstOrDefault();
                            if (text != null) text.Text = MapDingbats(text.Text, paragraph);
                            child.Attributes["style"].Remove();
                        }
                        if (child.Attributes["style"]?.Value == "font-family:\"ZapfDingbat Dop\"")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().FirstOrDefault();
                            if (text != null) text.Text = MapDingbatsDop(text.Text, paragraph);
                            child.Attributes["style"].Remove();
                        }
                        if (child.Attributes["style"]?.Value == "font-family:ZapfDingbats")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().Single();
                            if (!string.IsNullOrWhiteSpace(text.Text)) throw new Exception();
                            child.Attributes["style"].Remove();
                        }
                        if (child.Attributes["style"]?.Value == "font-family:\"Lucida Icons\"")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().FirstOrDefault();
                            if (text != null) text.Text = MapLucidaIcons(text.Text, paragraph);
                            child.Attributes["style"].Remove();
                        }
                        if (child.Attributes["style"]?.Value == "font-family:Antiqua")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().FirstOrDefault();
                            if (text != null) text.Text = MapAntiqua(text.Text, paragraph);
                            child.Attributes["style"].Remove();
                        }
                        
                        if (child.Attributes["style"]?.Value == "font-family:\"Lucida Bright Math Symbol\"")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().First();
                            text.Text = MapLucida(text.Text, paragraph);
                            child.Attributes["style"].Remove();
                        }
                        if (child.Attributes["style"]?.Value == "display:none;mso-hide:all")
                        {
                            var text = child.ChildNodes.OfType<HtmlTextNode>().First();
                            text.Text = $"(TODO: {text.Text})";
                            child.Attributes["style"].Remove();
                        }
                        if (child.HasAttributes && child.Name != "st1:metricconverter") throw new Exception();
                        
                        if (child.Name != "sup" || (child.FirstChild as HtmlTextNode)?.Text.Contains("♠") == true)
                        {
                            child = RemoveAndPromoteChildren(child);
                            continue;
                        }

                        break;
                    case "br":
                    case "b":
                    case "i":
                    case "#text":
                    case "#comment":
                        if (child.HasAttributes) throw new Exception();
                        break;
                    default:
                        throw new Exception("Unexpected tag: " + child.Name);
                }

                child = Next(paragraph, child);
            }
        }

        static void HandleAcc(HtmlNode child, string fontFamily, string combiningAccent)
        {
            if (child.Attributes["style"]?.Value == "font-family:" + fontFamily)
            {
                var text = child.ChildNodes.OfType<HtmlTextNode>().FirstOrDefault();
                if (text != null)
                    text.Text = text.Text.Insert(1, combiningAccent);
                child.Attributes["style"].Remove();
            }
        }

        static HtmlNode Next(HtmlNode node, HtmlNode child)
        {
            return child.FirstChild ?? child.NextSibling ?? NextSiblingOfParent(child, node);
        }

        static HtmlNode NextSiblingOfParent(HtmlNode child, HtmlNode paragraph)
        {
            if (child.ParentNode == paragraph) return null;
            HtmlNode parentNodeNextSibling() => child.ParentNode.NextSibling;
            while (parentNodeNextSibling() == null)
            {
                child = child.ParentNode;
                if (child.ParentNode == paragraph) return null;
            }
            return parentNodeNextSibling();
        }

        static HtmlNode RemoveAndPromoteChildren(HtmlNode child)
        {
            if (child.FirstChild == null)
            {
                child.ChildNodes.Add(child.OwnerDocument.CreateTextNode(""));
            }
            HtmlNode firstChild = child.FirstChild;
            HtmlNode parent = child.ParentNode;
            int i = parent.ChildNodes.IndexOf(child);
            child.Remove();
            foreach (var newChild in child.ChildNodes)
            {
                parent.ChildNodes.Insert(i, newChild);
                ++i;
            }
            return firstChild;
        }

        static bool RemoveStyle(string s)
        {
            int i = s.IndexOf(':');
            string name = s.Substring(0, i);
            switch (name)
            {
                case "mso-tab-count":
                case "font-size":
                case "mso-bidi-font-size":
                case "mso-bidi-font-style":
                case "position":
                case "top":
                case "mso-text-raise":
                case "letter-spacing":
                case "mso-bidi-font-weight":
                case "mso-ansi-language":
                case "mso-spacerun":
                case "mso-bidi-font-family":
                case "mso-fareast-font-family":
                case "mso-bookmark":
                    return true;
            }
            switch (s)
            {
                case "mso-bidi-font-weight:normal":
                case "mso-list:Ignore":
                case "font-family:TimesET":
                case "font-family:\"Courier New\"":
                //case "mso-bidi-font-family:\"Courier New\"":
                case "font-family:Symbol":
                case "font:7.0pt \"Times New Roman\"":
                case "background:red": // TODO: Что это за выделение?
                case "mso-highlight:red":
                case "text-decoration:none":
                    return true;
            }
            return false;
        }
    }
}
