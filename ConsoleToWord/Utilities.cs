using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleToWord;

public static class Utilities
{
    public static Paragraph CreateParagraph(string text, bool isBold = false, int fontSize = 12, JustificationValues justification = default)
    {
        Paragraph paragraph = new Paragraph();
        Run run = new Run();
        RunProperties runProperties = new RunProperties();

        if (isBold)
        {
            runProperties.Append(new Bold());
        }

        runProperties.Append(new FontSize() { Val = (fontSize * 2).ToString() });
        run.Append(runProperties);
        run.Append(new Text(text));

        ParagraphProperties paragraphProperties = new ParagraphProperties();
        paragraphProperties.Justification = new Justification() { Val = justification };
        paragraph.Append(paragraphProperties);

        paragraph.Append(run);
        return paragraph;
    }

    // Método auxiliar para criar parágrafos com parte do texto em negrito
    public static Paragraph CreateParagraphWithBold(string boldText, string normalText, JustificationValues justification = default)
    {
        Paragraph paragraph = new Paragraph();

        Run boldRun = CreateRun(boldText, isBold: true);
        Run normalRun = CreateRun(normalText);

        paragraph.Append(new ParagraphProperties(new Justification() { Val = justification }));
        paragraph.Append(boldRun);
        paragraph.Append(normalRun);

        return paragraph;
    }

    // Método auxiliar para criar corridas de texto
    public static Run CreateRun(string text, bool isBold = false, int fontSize = 12)
    {
        Run run = new Run();
        RunProperties runProperties = new RunProperties();

        if (isBold)
        {
            runProperties.Append(new Bold());
        }

        runProperties.Append(new FontSize() { Val = (fontSize * 2).ToString() });
        run.Append(runProperties);
        run.Append(new Text(text));

        return run;
    }
}