using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using NUnit.Framework;

namespace DocxToHtmlConverter
{
    [TestFixture]
    public class Tests
    {
        [Test]
        public void Test1()
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(@"<p class=MsoNormal><u><span lang=RU style='font-size:6.0pt;mso-bidi-font-size:
10.0pt;font-family:""ZapfDingbats BT""'>t</span></u><u><span lang=RU
style='font-family:TimesET'><span style='mso-tab-count:1'>                     </span><b
style='mso-bidi-font-weight:normal'>пацанв</b></span></u><b style='mso-bidi-font-weight:
normal'><u><span lang=RU style='font-family:Tim_acc'>а</span></u></b><u><span
lang=RU style='font-family:TimesET'><span style='mso-tab-count:1'>      </span>ж<span
style='mso-tab-count:1'>             </span>1b— (<i style='mso-bidi-font-style:
normal'>собир.</i>)<o:p></o:p></span></u></p>
");
            Assert.AreEqual("♠ <b>пацанва́</b> ж 1b— (<i>собир.</i>)", Program.ToText(doc.DocumentNode.FirstChild));
        }

        [Test]
        public void Test2()
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(@"<b><sup style='font-family:""ZapfDingbats BT""'>t</sup><sup> 1-4</sup>ла́ва</b> ж 1a");
            Assert.AreEqual("♠<sup> 1-4</sup>ла́ва", Program.ToText(doc.DocumentNode.FirstChild));
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ConvertNames();
            ConvertMainDictionary();
        }

        static void ConvertNames()
        {
            var names = File.ReadLines(@"..\..\names.txt")
                .Except(new []
                {
                    "(Вели́кий) Но́вгород (п +) м 1а",
                })
                .Select(line => line.Replace("мо+", "мо⁺"))
                .Select(line => line.Replace("жо+", "жо⁺"))
                .Select(line => line.Replace(" см. ", " <i>см.</i> "))
                .Select(line => line.Replace("(река)", "(<i>река</i>)"))
                .Select(line => line.Replace("(страна)", "(<i>страна</i>)"))
                .Select(line => line.Replace("(гопокор.)", "(<i>гопокор.</i>)"))
                .Select(ParseName)
                .Select(ToLine);
                
            File.WriteAllLines(@"C:\Code\gramdict.ru.api\ZalizniakDictionary\names.txt", names);
        }

        static void ConvertMainDictionary()
        {
            string path = @"..\..\..\..\docxconverter";

            string filePath = Path.Combine(path, "all2.htm");

            const string outputPath = "zal.txt";
            var fileInfo = new FileInfo(outputPath);
            if (!fileInfo.Exists || fileInfo.LastWriteTimeUtc < new FileInfo(filePath).LastWriteTimeUtc)
                CleanDom(filePath, outputPath);

            string text = File.ReadAllText(outputPath);

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
            text = text.Replace("дать</b> св △ b/c':́ <i>", "дать</b> св △ b/c': <i>");

            File.WriteAllText("all.txt", text);

            string[] lines = text.Split(new[] {Environment.NewLine}, StringSplitOptions.None);

            string finalOutputPath = @"C:\Code\gramdict.ru.api\ZalizniakDictionary\zaliznyak.txt";

            File.WriteAllLines(finalOutputPath, lines
                .Where(line => !line.Contains("TODO:") && !string.IsNullOrWhiteSpace(line))
                .Where(line => !line.StartsWith("<b>споко́н</b>"))
                .Where(line => !line.StartsWith("<b>угль</b>"))
                .Where(line => !line.StartsWith("<b>жураве́ль</b>"))
                .Where(line => !line.StartsWith("<b>огнь</b>"))
                .Where(line => !line.StartsWith("<b>ви́хорь</b>"))
                .Select(Parse)
                .Select(ToLine));
        }

        public static string ToLine(Entry e)
        {
            return $"{e.Lemma}|{e.Symbol}" + (string.IsNullOrEmpty(e.Grammar) ? "" : $"|{e.Grammar}");
        }

        static Entry Parse(string line, int number)
        {
            Match match = Regex.Match(line, @"^\s*(?<spade>♠*)\s*<b>(<sup>[ -0123456789]+</sup>)?(?<lemma>[- а-яА-ЯёЁ\u0300\u0301]+:?)</b>\s*(?<rest>.+)$");
            if (!match.Success) throw new Exception($"No match: ({number+1}) {line}");
            string parensSymbolGrammar = match.Groups["rest"].Value;
            (string symbolGrammar, string parens) = SplitParens(parensSymbolGrammar);
            string fullestSymbolsRegex = GetFullestSymbolsRegex();
            string symbol;
            string grammar;
            if (match.Groups["lemma"].Value.EndsWith(":"))
            {
                symbol = "";
                grammar = symbolGrammar;
            }
            else
            {
                Match match1 = Regex.Match(symbolGrammar, $@"^\s*(?<symbol>{fullestSymbolsRegex}(,|:|;)?)(\s+(?<grammar>.*))?$");
                if (!match1.Success) throw new Exception("Regex failed: " + line);
                symbol = match1.Groups["symbol"].Value.TrimEnd();
                grammar = match1.Groups["grammar"].Value;
            }
            return new Entry
            {
                Lemma = match.Groups["lemma"].Value,
                Symbol = symbol,
                Grammar = grammar
            };
        }

        public static Entry ParseName(string line, int number)
        {
            string fullestSymbolsRegex = GetFullestSymbolsRegex();
            Match match = Regex.Match(line, $@"^(?<lemma>[-’ а-яА-ЯёЁ\u0300\u0301\(\)]+:?)\s+(?<symbol>{fullestSymbolsRegex}(,|:|;)?)(\s+(?<grammar>.+))?$");
            if (!match.Success) throw new Exception($"No match: ({number+1}) {line}");
            return new Entry
            {
                Lemma = match.Groups["lemma"].Value,
                Symbol = match.Groups["symbol"].Value,
                Grammar = match.Groups["grammar"].Value
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
            string[] symbols =
                "ф.,мо⁺,жо⁺,м,мо,ж,жо,с,со,жо,мо-жо,мн.,мн. одуш.,мн. неод.,мн. _от_,п,мс,мс-п,числ.,числ.-п,св,нсв,св-нсв,св/нсв,н,част.,част.(_усилительная_),союз,предл.,предик.,вводн.,межд.,сравн.,§1,§2,предикативное мс,_см._"
                    .Split(',')
                    .OrderByDescending(s => s.Length).ToArray();
            // TODO нп, безл., многокр.

            string symbolsRegex = RegexEscape(symbols);

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

        public class Entry
        {
            public string Lemma, Symbol, Grammar;
        }
        
        static void CleanDom(string htmlInputPath, string outputPath)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlInputPath, Encoding.UTF8);
            HtmlNodeCollection entries = doc.DocumentNode.SelectNodes("html/body/div/p");
            Console.WriteLine(entries.Count);
            IEnumerable<string> lines = entries
                .Skip(8)
                .Select(ToText)
                .Where(line => line.Length > 8); // skip letter headings
            File.WriteAllLines(outputPath, lines);
        }

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
