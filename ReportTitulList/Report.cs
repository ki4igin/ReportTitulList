using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ReportTitulList;

public class Report : IDisposable
{
    private const string StyleId = "MyStyle";
    private const string StyleIdTable = "TableStyle";

    private readonly WordprocessingDocument _document;
    private readonly Body _body;

    private int _pageCount;

    public Report(string fileName)
    {
        _document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);

        MainDocumentPart mainPart = _document.AddMainDocumentPart();
        mainPart.Document = new();
        _body = mainPart.Document.AppendChild(new Body());

        _body.Append(new SectionProperties(new PageMargin {Left = 850, Right = 850})); // 1.5 см


        StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();


        Styles styles = new(
            new Style(
                new Name {Val = StyleId},
                new BasedOn {Val = "Normal"},
                new ParagraphProperties(
                    new SpacingBetweenLines
                        {Line = "360", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0"},
                    new Indentation {Left = "0", Right = "0"},
                    new Justification {Val = JustificationValues.Center},
                    new RunProperties(
                        new RunFonts {Ascii = "Times New Roman", HighAnsi = "Times New Roman"},
                        new FontSize {Val = "28"} // 14 font
                    ))
            ),
            new Style(
                new Name {Val = StyleIdTable},
                new BasedOn {Val = "Normal"},
                new ParagraphProperties(
                    new SpacingBetweenLines
                        {Line = "240", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0"},
                    new Indentation {Left = "0", Right = "0"},
                    new Justification {Val = JustificationValues.Center},
                    new RunProperties(
                        new RunFonts {Ascii = "Times New Roman", HighAnsi = "Times New Roman"},
                        new FontSize {Val = "24"} // 12 font
                    ))
            )
        );

        stylePart.Styles = styles;
    }

    public void Add(ReportSettings settings)
    {
        // New Page
        if (_pageCount++ > 0)
            _body.Append(new Paragraph(new Run(new Break() {Type = BreakValues.Page})));

        string workName = settings.WorkType;
        int workCount = settings.WorkCount;

        List<string> fioList = settings.Names.Select(ConvertFullNameToShortName).ToList();
        int workAllCount = fioList.Count * workCount;

        // Add Styled Text
        Ext.Style = StyleId;
        _body
            .AddParagraphToBody($"{settings.Year}, {settings.Semester} семестр")
            .AddParagraphToBody($"{settings.Group}")
            .AddParagraphToBody($"«{settings.Discipline}»")
            .AddParagraphToBody($"{workAllCount} шт.")
            .AddParagraphToBody("");

        string[][] tableContent = new string[fioList.Count + 1][];
        tableContent[0] = new[] {"№", "ФИО"}
            .Concat(Enumerable.Range(1, workCount).Select(s => $"{workName}{s}")).ToArray();
        for (int i = 1; i <= fioList.Count; i++)
        {
            tableContent[i] = new[] {$"{i}", fioList[i - 1]}
                .Concat(Enumerable.Range(1, workCount).Select(_ => "+")).ToArray();
        }

        Table table = Ext.CreateTable(tableContent, StyleIdTable);
        _body.Append(table);
    }

    public void Dispose()
    {
        _document.Dispose();
        GC.SuppressFinalize(this);
    }

    private static string ConvertFullNameToShortName(string fullName)
    {
        string[] fio = fullName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return fio.Length switch
        {
            1 => fio[0],
            2 => $"{fio[0]} {fio[1].First()}.",
            _ => $"{fio[0]} {fio[1].First()}.{fio[2].First()}."
        };
    }
}