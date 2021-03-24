using NUnit.Framework;

namespace DocxToHtmlConverter.Tests
{
    [TestFixture]
    public class LemmaExtensionsTests
    {
        [Test]
        [TestCase("кот", "кот 0")]
        [TestCase("рыба́к", "рыбак 2")]
        [TestCase("учёба", "учеба 0,3")]
        [TestCase("бельё", "белье 0,1")]
        [TestCase("Кзы́л-Орда́", "Кзыл-Орда 2+1")]
        [TestCase("Золота́я Орда́", "Золотая Орда 2+1")]
        [TestCase("Москва́-река́", "Москва-река 1+1")]
        [TestCase("ну́-ка", "ну-ка 1+0")]
        [TestCase("флё̀рдора́нж", "флердоранж 3,8.3")]
        [TestCase("пти́чка-невели́чка", "птичка-невеличка 4+4")]
        [TestCase("д’Артанья́н", "д’Артаньян 2")]
        [TestCase("Д’Анну́нцио", "Д’Аннунцио 5")]
        [TestCase("Эль Гре́ко", "Эль Греко 0+3")]
        [TestCase("Э́льба", "Эльба 5")]
        [TestCase("и́на́че", "иначе 3//5")]
        public void StripStressMarksTest3(string lemma, string expected)
        {
            Assert.AreEqual(expected, lemma.ConvertStressMarksToNumbers());
        }
    }
}