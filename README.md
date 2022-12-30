# docx2html

Этот репозиторий содержит:

1. Файл all.html, полученный из файла all.docx функцией "Save as HTML" в Microsoft Word.
2. Файл names.txt, полученный OCR-распознаванием Приложения «Имена собственные» к 6-му изданию словаря Зализняка с последующими ручными правками.
3. Консольную программу на языке C#, которая преобразует эти два файла в html/txt, [исправляя ошибки](https://github.com/gramdict/docx2html/blob/master/DocxToHtmlConverter/Program.CorrectHtml.cs#L7) в all.html.

## Licensing

1. all.html, names.txt: CC-BY-NC
2. All other files: MIT
