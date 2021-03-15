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
    static class Program
    {
        static void Main(string [] args)
        {
            var stopwatch = Stopwatch.StartNew();

            File.WriteAllLines(args[1],
                ConvertNames(File.ReadLines(@"..\..\..\names.txt")).Select(ToOutputFormat));

            File.WriteAllLines(args[0], 
                ConvertCommonPart(GetCleanHtml(@"..\..\..\all.html")).Select(ToOutputFormat));
            
            Console.WriteLine($"{stopwatch.Elapsed.TotalMilliseconds:N0} ms");
        }

        private static IEnumerable<Entry> ConvertNames(IEnumerable<string> lines) =>
            lines
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

        private static string CorrectHtml(string text)
        {
            text = text.Replace("<![if !supportLists]>t <![endif]>", "");
            text = text.Replace("­", ""); // знак переноса
            text = text.Replace("<b>◑</b>", "◑");
            text = text.Replace("дидакти́чеcкий", "дидакти́ческий"); // латинская с
            text = text.Replace("cта́ксель", "ста́ксель"); // латинская с
            text = Regex.Replace(text, "\u00a0+", " ");
            //text = Regex.Replace(text, "\\s+", " ");
            text = text.Replace("</b><b>", "");
            text = text.Replace("</b>-<b>", "-");
            text = text.Replace("</sup><sup>", "");
            text = text.Replace("<sup> </sup>", "");
            text = text.Replace("<b>♠", "♠<b>");
            text = text.Replace("<b> ", " <b>");
            text = text.Replace("<b>♠", "♠<b>");
            text = text.Replace(" </b>", "</b> ");
            text = text.Replace("\n <b>", "\n<b>");
            text = text.Replace("<i> ", " <i>");
            text = text.Replace(" </i>", "</i> ");
            text = Regex.Replace(text, "</i>([,. \\(\\)«»;-]*)<i>", "$1");
            text = Regex.Replace(text, "<i>\\s*</i>", "");
            text = Regex.Replace(text, "<b>\\s*</b>", "");
            text = Regex.Replace(text, @"\s+\r\n", "\r\n");
            text = text.Replace("§ ", "§");
            text = text.Replace("// ", "//");
            text = Regex.Replace(text, "(?<!&gt);\r\n", ";<br/>");
            text = text.Replace("<i>сравн.</i> ши́ре\r\n", "<i>сравн.</i> ши́ре<br/>");
            text = text.Replace("на колени</i>)\r\n", "на колени</i>)<br/>");
            text = text.Replace("(отмаха́ть)\r\n нсв", "(отмаха́ть)<br/>нсв");
            text = text.Replace("летая, побывать всюду</i>);\r\n", "летая, побывать всюду</i>);<br/>");
            text = text.Replace("\r\n св (<i>дать трещину", "<br/> св (<i>дать трещину");
            text = text.Replace("<i>(", "(<i>");
            text = text.Replace(")</i>", "</i>)");
            text = text.Replace("<b>лис(с)або́нский</b> п 3a✕~", "<b>лиссабо́нский</b> п 3a✕~ [//<b>лисабо́нский</b>]");
            text = text.Replace("<b>ïñî́ó</b> ñ 0", "<b>пс́оу</b> с 0");
            text = text.Replace("<b>лис(с)або́нец</b> мо 5*a", "<b>лиссабо́нец</b> мо 5*a [//<b>лисабо́нец</b>]");
            text = text.Replace("<sup> <b>2</b></sup><b>", "<b><sup> 2</sup>");
            text = text.Replace("без удар</i>.", "без удар.</i>");
            text = text.Replace("<i>нормально без удар</i>", "<i>нормально без удар.</i>");
            text = text.Replace("<i>нормально  без удар.</i>", "<i>нормально без удар.</i>");
            text = text.Replace("<i>в знач</i>.", "<i>в знач.</i>");
            text = text.Replace("(<i>в знач. «иди») повел. от</i>", "(<i>в знач. «иди»</i>) <i>повел. от</i>");
            text = text.Replace("</b>:", ":</b>");
            text = text.Replace("иват</b>ь", "ивать</b>");
            text = text.Replace("сн-нсв", "св-нсв");
            text = Regex.Replace(text, @"(\d)" + Regex.Escape("<sup>*</sup>"), "$1°");
            text = text.Replace("дать</b> св △ b/c':́ <i>", "дать</b> св △ b/c': <i>");
            text = text.Replace("✧ ́за́", "✧ за́"); // лишний знак ударения в статье город
            text = text.Replace(";<br/><b>", "\r\n<b>");
            text = text.Replace("о<i>т</i>", "<i>от</i>");
            text = text.Replace("пс́оу", "псо́у");
            text = text.Replace("<b>из-за</b>", "<b>и́з-за</b>");
            text = text.Replace("<b>из-под</b>", "<b>и́з-под</b>");
            text = text.Replace("<b>из-подо</b>", "<b>и́з-подо</b>");
            text = text.Replace("<b>по-за</b>", "<b>по́-за</b>");
            text = text.Replace("<b>по-над</b>", "<b>по́-над</b>");
            text = text.Replace("<b>подо</b>", "<b>по́до</b>");
            text = text.Replace("<b>надо</b>", "<b>на́до</b>");
            text = text.Replace("<b>безо</b>", "<b>бе́зо</b>");
            text = text.Replace("<b>предо</b>", "<b>пре́до</b>");
            text = text.Replace("<b>ото</b>", "<b>о́то</b>");
            text = text.Replace("<b>изо</b>", "<b>и́зо</b>");
            text = text.Replace("сужден́о", "суждено́");
            text = text.Replace("<b>аир</b>", "<b>а́ир</b>");
            text = text.Replace("<b>бородинский</b>", "<b>бороди́нский</b>");
            text = text.Replace("<b>дрочона</b>", "<b>дрочо́на</b>");
            text = text.Replace("<b>лысенковский</b>", "<b>лысе́нковский</b>");
            text = text.Replace("<b>шлиссельбуржец</b>", "<b>шлиссельбу́ржец</b>");
            text = text.Replace("<b>подслащивать</b>", "<b>подсла́щивать</b>");
            text = text.Replace("<b>приспускать</b>", "<b>приспуска́ть</b>");
            text = text.Replace("<b>н́ельмовый</b>", "<b>не́льмовый</b>");
            text = text.Replace("<b>т́естовый</b>", "<b>те́стовый</b>");
            text = text.Replace("<b>капр́изник</b>", "<b>капри́зник</b>");
            text = text.Replace("<b>краснодер́евщик</b>", "<b>краснодере́вщик</b>");
            text = text.Replace("<b>желтол́озник</b>", "<b>желтоло́зник</b>");
            text = text.Replace("<b>придор́ожник</b>", "<b>придоро́жник</b>");
            text = text.Replace("<b>железнодор́ожник</b>", "<b>железнодоро́жник</b>");
            text = text.Replace("<b>неприме́тн́ость</b>", "<b>неприме́тность</b>");
            text = text.Replace("<b>ѓикнуть</b>", "<b>ги́кнуть</b>");
            text = text.Replace("<b>чил́икнуть</b>", "<b>чили́кнуть</b>");
            text = text.Replace("<b>ѓалицкий</b>", "<b>га́лицкий</b>");
            text = text.Replace("<b>перебега́тъ</b>", "<b>перебега́ть</b>");
            text = text.Replace("<b>заса́харивать-</b>", "<b>заса́харивать</b>");
            text = text.Replace("долев́ой", "долево́й");
            text = text.Replace("́ уже", "у́же");
            text = text.Replace("<i> ", " <i>");
            text = text.Replace(" </i>", "</i> ");
            text = text.Replace("</i>.", ".</i>");
            text = text.Replace("мо-жо (<i>бранно о человеке</i>)", "мо-жо 1a (<i>бранно о человеке</i>)");
            text = text.Replace("<i>Р. мн. затрудн. (бранно о человеке</i>)",
                "<i>Р. мн. затрудн.</i> (<i>бранно о человеке</i>)");
            text = text.Replace("Кул́игин", "Кули́гин");
            text = text.Replace("в́скри́кнуть", "вскри́кнуть");
            text = text.Replace("зу́б́ча́тый", "зу́бча́тый");
            text = text.Replace("<i>кф (затрудн.</i>)", "<i>кф (затрудн.)</i>");

            text = Regex.Replace(text, @"<i>Р\. мн\. затрудн\. \((.+)</i>\)", "<i>Р. мн. затрудн.</i> (<i>$1</i>)");
            text = Regex.Replace(text, ";<br/>\\s*♠?\\s*<b>", "\r\n♠ <b>");
            return text;
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
                def => $"{def.Symbol}" + (string.IsNullOrEmpty(def.Grammar) ? "" : $"|{def.Grammar}")));

        static Entry Parse(string line, int number)
        {
            Match match = Regex.Match(line, @"^\s*(?<spade>♠*)\s*<b>(<sup>[ -0123456789]+</sup>)?(?<lemma>[- а-яА-ЯёЁ\u0300\u0301]+:?)</b>\s*(?<rest>.+)$");
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
                definitions = symbolGrammar.Split("<br/>")
                    .Select(def => ParseDefinition(line, def))
                    .ToList();
            }
            return new Entry
            {
                Lemma = match.Groups["lemma"].Value,
                Definitions = definitions
            };
        }

        private static Definition ParseDefinition(string line, string symbolGrammar)
        {
            Match match1 = Regex.Match(symbolGrammar, $@"^♠?\s*(?<symbol>{FullestSymbolsRegex}(,|:|;)?)(\s+(?<grammar>.*))?$");
            if (!match1.Success) throw new Exception("Regex failed: " + line);
            var def = new Definition
            {
                Symbol = match1.Groups["symbol"].Value.TrimEnd(),
                Grammar = match1.Groups["grammar"].Value
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
                case "<i>нормально без удар.</i>":
                case "<i>часто без удар.</i>: еще":
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
            return Regex.Replace(s, "_", m => ++i % 2 == 1 ? "<i>" : "</i>");
        }

        public class Definition
        {
            public string Symbol, Grammar;
        }
        
        public class Entry
        {
            public string Lemma;
            public IEnumerable<Definition> Definitions;
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
