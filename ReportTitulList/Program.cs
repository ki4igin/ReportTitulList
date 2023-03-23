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

// Создание стиля абзаца

StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
string styleId = "MyStyle";

Styles styles = new(
    new Style(
        new Name { Val = styleId },
        new BasedOn { Val = "Normal" },
        new ParagraphProperties(
            new SpacingBetweenLines { Line = "360", LineRule = LineSpacingRuleValues.Auto, Before = "0", After = "0" },
            new Indentation { Left = "0", Right = "0" },
            new Justification { Val = JustificationValues.Center },
            new RunProperties(
                new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                new FontSize { Val = "28" } // 14 шрифт
            ))
    ));

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


Table table = new();

TableProperties tableProperties = new(
    new TableBorders(
        new TopBorder { Val = new(BorderValues.Single), Size = 2 },
        new BottomBorder { Val = new(BorderValues.Single), Size = 2 },
        new LeftBorder { Val = new(BorderValues.Single), Size = 2 },
        new RightBorder { Val = new(BorderValues.Single), Size = 2 },
        new InsideHorizontalBorder { Val = new(BorderValues.Single), Size = 2 },
        new InsideVerticalBorder { Val = new(BorderValues.Single), Size = 2 }
    ),
    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
);
table.AppendChild(tableProperties);


TableRow titleRow = new() { TableRowProperties = new(new TableRowHeight { Val = 4540 }) }; // 8 см
titleRow.Append(new TableCell(
    Ext.CreateParagraph("ФИО"),
    new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center })));
for (int i = 0; i < workCount; i++)
{
    titleRow.Append(new TableCell(new List<OpenXmlElement>()
        .Append(Ext.CreateParagraph($"{workName}{i + 1}"))));
}

table.Append(titleRow);

for (int i = 0; i < fioList.Count; i++)
{
    TableRow row = new();
    row.Append(new TableCell(Ext.CreateParagraph($"{i + 1}"))
    { TableCellProperties = new() { TableCellVerticalAlignment = new() { Val = TableVerticalAlignmentValues.Center } } });
    row.Append(new TableCell(Ext.CreateParagraph($"{fioList[i]}")));
    for (int j = 0; j < workCount; j++)
    {
        row.Append(new TableCell(Ext.CreateParagraph("+")));
    }

    table.Append(row);
}

body.AddParagraphToBody("");


body.Append(table);


public static class Ext
{
    public static string Style { get; set; }

    public static Body AddParagraphToBody(this Body body, string text)
    {
        Paragraph paragraph = new()
        {
            ParagraphProperties = new(new ParagraphStyleId { Val = Style })
        };
        paragraph.Append(new Run(new Text(text)));
        body.Append(paragraph);
        return body;
    }

    public static Paragraph CreateParagraph(string text)
    {
        Paragraph paragraph = new()
        {
            ParagraphProperties = new(new ParagraphStyleId { Val = Style })
        };
        paragraph.Append(new Run(new Text(text)));
        return paragraph;
    }
}