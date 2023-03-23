using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;

Console.WriteLine("Hello, World!");

string fileName = "example.docx";

// Создание документа
using WordprocessingDocument document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);
// Создание основной части документа
MainDocumentPart mainPart = document.AddMainDocumentPart();
mainPart.Document = new();
Body body = mainPart.Document.AppendChild(new Body());
body.Append(new SectionProperties(new PageMargin {Left = 850, Right = 850})); // 1.5 см
// Создание стиля абзаца

StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
string styleId = "MyStyle";
string styleIdTable = "TableStyle";

Styles styles = new(
    new Style(
        new Name {Val = styleId},
        new BasedOn {Val = "Normal"},
        new ParagraphProperties(
            new SpacingBetweenLines {Line = "360", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0"},
            new Indentation {Left = "0", Right = "0"},
            new Justification {Val = JustificationValues.Center},
            new RunProperties(
                new RunFonts {Ascii = "Times New Roman", HighAnsi = "Times New Roman"},
                new FontSize {Val = "28"} // 14 шрифт
            ))
    ),
    new Style(
        new Name {Val = styleIdTable},
        new BasedOn {Val = "Normal"},
        new ParagraphProperties(
            new SpacingBetweenLines {Line = "240", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0"},
            new Indentation {Left = "0", Right = "0"},
            new Justification {Val = JustificationValues.Center},
            new RunProperties(
                new RunFonts {Ascii = "Times New Roman", HighAnsi = "Times New Roman"},
                new FontSize {Val = "24"} // 12 шрифт
            ))
    )
);

stylePart.Styles = styles;

string workName = "ЛР";
int workCount = 4;
List<string> workList = new();
for (int i = 0; i < workCount; i++)
{
    workList.Add($"{workName}{i + 1}");
}

string str;
int cnt = 0;
List<string> names = new();
while (cnt < 2)
{
    str = Console.ReadLine()!;
    if (str.Length > 3)
    {
        names.Add(str);
        cnt = 0;
    }
    else
    {
        cnt++;
    }
}

int workAllCount = names.Count * workCount;

List<string> fioList = names
    .Select(s => $"{s.Split().ToArray()[0]} {s.Split().ToArray()[1][0]}.{s.Split().ToArray()[2][0]}.").ToList();
Console.WriteLine(string.Join("\n", fioList));
Ext.Style = styleId;


// Добавление строк с применением стиля
body
    .AddParagraphToBody("2022-2023")
    .AddParagraphToBody("1 семестр")
    .AddParagraphToBody("«Устройства на основе ПЛИС»")
    .AddParagraphToBody("Кичигин А.А.")
    .AddParagraphToBody("СМ5-71")
    .AddParagraphToBody(string.Join(" + ", workList))
    .AddParagraphToBody($"{workAllCount}");


body.AddParagraphToBody("");

string[][] tableContent = new string[fioList.Count + 1][];
tableContent[0] = new[] {"№", "ФИО"}.Concat(Enumerable.Range(1, workCount).Select(s => $"{workName}{s}")).ToArray();
for (int i = 1; i <= fioList.Count; i++)
{
    tableContent[i] = new[] {$"{i}", fioList[i - 1]}.Concat(Enumerable.Range(1, workCount).Select(_ => "+")).ToArray();
}

Table table = Ext.CreateTable(tableContent, styleIdTable);
body.Append(table);


public static class Ext
{
    public static string Style { get; set; }

    public static Table CreateTable(string[][] content, string styleId)
    {
        Table table = new();

        TableProperties tableProperties = new(
            new TableBorders(
                new TopBorder {Val = new(BorderValues.Single), Size = 2},
                new BottomBorder {Val = new(BorderValues.Single), Size = 2},
                new LeftBorder {Val = new(BorderValues.Single), Size = 2},
                new RightBorder {Val = new(BorderValues.Single), Size = 2},
                new InsideHorizontalBorder {Val = new(BorderValues.Single), Size = 2},
                new InsideVerticalBorder {Val = new(BorderValues.Single), Size = 2}
            ),
            new TableWidth {Width = "5000", Type = TableWidthUnitValues.Pct},
            new TableCellVerticalAlignment {Val = TableVerticalAlignmentValues.Center}
        );
        table.AppendChild(tableProperties);

        foreach (string[] rowContent in content)
        {
            TableRow row = new() {TableRowProperties = new(new TableRowHeight {Val = 454})}; // 8 см
            foreach (string cellContent in rowContent)
            {
                row.Append(new TableCell(
                    new TableCellProperties(new TableCellVerticalAlignment {Val = TableVerticalAlignmentValues.Center}),
                    CreateParagraph(cellContent, styleId)
                ));
            }

            table.Append(row);
        }

        return table;
    }

    public static Body AddParagraphToBody(this Body body, string text)
    {
        Paragraph paragraph = new()
        {
            ParagraphProperties = new(new ParagraphStyleId {Val = Style})
        };
        paragraph.Append(new Run(new Text(text)));
        body.Append(paragraph);
        return body;
    }

    public static Paragraph CreateParagraph(string text) =>
        CreateParagraph(text, Style);

    public static Paragraph CreateParagraph(string text, string style)
    {
        Paragraph paragraph = new()
        {
            ParagraphProperties = new(new ParagraphStyleId {Val = style})
        };
        paragraph.Append(new Run(new Text(text)));
        return paragraph;
    }
}