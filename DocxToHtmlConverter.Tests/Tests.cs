using HtmlAgilityPack;
using NUnit.Framework;

namespace DocxToHtmlConverter.Tests
{
    [TestFixture]
    public class EndToEndTests
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

        [Test]
        [TestCase("✧ ́зу́б", "✧ зу́б")]
        public void RemoveStressMarksOverNonVowelsTest(string input, string expectedOutput)
        {
            Assert.AreEqual(expectedOutput, Program.RemoveStressMarksOverNonVowels(input));
        }
    }
}