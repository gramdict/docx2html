using System.Text.RegularExpressions;

namespace DocxToHtmlConverter
{
    public static partial class Program
    {
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
            text = Regex.Replace(text,
                @"\(<i>([^<]+)\), ",
                @"(<i>$1</i>), <i>");
            text = Regex.Replace(text,
                @" \(([^<)]+)</i>\)",
                "</i> (<i>$1</i>)");
            text = text.Replace("словам</i>и:", "словами</i>:");

            text = Regex.Replace(text, @"<i>Р\. мн\. затрудн\. \((.+)</i>\)", "<i>Р. мн. затрудн.</i> (<i>$1</i>)");
            text = Regex.Replace(text, ";<br/>\\s*♠?\\s*<b>", "\r\n♠ <b>");
            text = text.Replace("<i>, ", ", <i>");
            text = text.Replace("за́ два;́ на́ два", "за́ два; на́ два");
            text = text.Replace("фо̀тоискусство", "фо̀тоиску́сство");
            text = text.Replace("<i>́</i>", "");
            text = Regex.Replace(text, @"(\d)\(<i>([^)]+)</i>\)", "$1($2)");
            text = text.Replace("вы́ел; -а;", "вы́ел, -а;");
            text = text.Replace(" -</i>", "</i> -");
            text = text.Replace(
                "(<i>этот); не смешивать с прич. страд. от</i> ",
                "(<i>этот</i>); <i>не смешивать с прич. страд. от</i> ");
            text = text.Replace(
                "(<i>определенный); не смешивать с прич. страд. от</i> ",
                "(<i>определенный</i>); <i>не смешивать с прич. страд. от</i> ");
            text = text.Replace("(-д́ить)", "(-ди́ть)");
            text = text.Replace("(-с́ить)", "(-си́ть)");
            text = text.Replace("(-с́тить)", "(-сти́ть)");
            text = text.Replace("дожин́ать)", "дожина́ть)");
            text = text.Replace("пере́вёр́ты́вать", "перевёртывать");
            text = text.Replace("перево́ра́чивать", "перевора́чивать");
            text = text.Replace("раз́вёр́ты́вать", "развёртывать");
            text = text.Replace("вы́про́ки́ну́ть́ся", "вы́прокинуться");
            text = text.Replace("п́од́ъехать", "подъе́хать");
            text = RemoveStressMarksOverNonVowels(text);
            text = text.Replace("занумеро́вы́вать", "занумеро́вывать");
            text = text.Replace("пе́ре́ну́ме́ро́вывать", "перенумеро́вывать");
            text = text.Replace("пронумеро́вы́вать", "пронумеро́вывать");
            text = text.Replace("заноме́ро́вывать", "заномеро́вывать");
            text = text.Replace("пе́ре́номеро́вывать", "переномеро́вывать");
            text = text.Replace("про́но́ме́ро́вы́вать", "прономеро́вывать");
            text = text.Replace("<sup></sup>", "");
            text = text.Replace("<b></b>", "");
            text = text.Replace("</i><i>", "");
            text = text.Replace("</b><b>", "");
            text = text.Replace("мещё́ра", "мещёра");
            text = Regex.Replace(text, @"\(<i>([^<)]+)\); (?!\<)", "(<i>$1</i>); <i>");
            text = text.Replace("<b>]</b>", "]");
            text = text.Replace("<b>◑</b>", "◑");
            text = text.Replace("љ", "");
            text = text.Replace("с<i>пряж", "<i>спряж");
            text = text.Replace("(<i>сустав)</i>", "(<i>сустав</i>)");
            text = text.Replace("(<i>см.) и</i>", "(<i>см.</i>) <i>и</i>");
            text = Regex.Replace(text, @"c (\d)", "с $1");
            text = text.Replace("(<i>о цене) и</i>", "(<i>о цене</i>) <i>и</i>");
            text = text.Replace("<i>повел. нет,</i>", "<i>повел. нет</i>,");
            text = text.Replace("△а", "△a");
            text = text.Replace("<i>прич. прош. нет (предположить; признать</i>",
                "<i>прич. прош. нет</i> (<i>предположить; признать</i>");
            text = text.Replace("<i>,</i>", ",");
            text = Regex.Replace(text, ",([^ <])", ", $1"); // add space after comma
            text = text.Replace("</i>»", "»</i>");
            text = text.Replace("св △ ", "св △");
            text = text.Replace("нп △ ", "нп △");
            text = text.Replace("<i>в св) и</i>", "<i>в св</i>) <i>и</i>");
            text = Regex.Replace(text, @"◑(.)\(<i>([^<]*)</i>\)", "◑$1($2)");
            text = text.Replace("|плется //щи́пется", "|плется //-пется");
            text = text.Replace("</i> — <i>", " — ");
            text = text.Replace("), спряж. см.</i>", "</i>), <i>спряж. см.</i>");
            text = text.Replace("заг|оню́, -онит", "заг|оню́, -о́нит");
            text = text.Replace("a '", "a'");
            text = text.Replace(
                "<b>отсня́ться</b> св 14с/c''",
                "<b>отсня́ться</b> св 14с/c'' (-им-)");
            text = text.Replace(
                "струга́ться</b>] ◑1(-а́-)",
                "струга́ться</b>] ◑I(-а́-)");
            text = text.Replace(
                "<b>уня́ться</b> св 14b/c ",
                "<b>уня́ться</b> св 14b/c'' ");
            text = text.Replace(
                "погрести́</b> св 7b/b (-б-), ё (<i>грести некоторое время</i>)",
                "погрести́</b> св нп 7b/b (-б-), ё (<i>грести некоторое время</i>)");
            text = text.Replace(
                "<b>гнести́</b> нсв 7b/b (-т-)⑨",
                "<b>гнести́</b> нсв 7b/b (-т-)");
            text = text.Replace(" //", "//");
            text = text.Replace(
                "<b>сжева́ть</b> св 2b, ё ◑1",
                "<b>сжева́ть</b> св 2b, ё ◑I"
            );
            text = text.Replace(
                "уволочи́ть</b>] ◑1(-а-)",
                "уволочи́ть</b>] ◑I(-а-)");
            text = text.Replace(
                "<b>отчленя́ть</b> св",
                "<b>отчленя́ть</b> нсв");
            text = text.Replace("ж 8з", "ж 8a");
            text = text.Replace("пешо́чком</b>] (", "пешо́чком</b> (");
            text = text.Replace("(<i>бежать, минуя что-л.</i> ◑5", "(<i>бежать, минуя что-л.</i>) ◑5");
            text = text.Replace("◑I(-а-)", "◑I(-а́-)");
            text = text.Replace("), Р.</i>", "</i>), <i>Р.</i>");
            text = text.Replace("(-c-)", "(-с-)");
            text = text.Replace("4с", "4c");
            text = text.Replace("lа", "1a");
            text = text.Replace("1а", "1a");
            text = text.Replace(
                "(дать сигнал флагом; кончить махать)",
                " (<i>дать сигнал флагом; кончить махать</i>)");
            text = text.Replace("созрев<i>а</i>ть", "созревать");
            text = text.Replace("cплю́щивать", "сплю́щивать");
            text = text.Replace("созрев</i>а<i>ть", "созревать");
            text = text.Replace("<i>В.=И.</i>)", "<i>В.=И.</i>");
            text = text.Replace("<b>чихну́ть</b> св нп 3b✕", "<b>чихну́ть</b> св нп 3b");
            return text;
        }
    }
}