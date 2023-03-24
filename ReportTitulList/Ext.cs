using DocumentFormat.OpenXml.Wordprocessing;

namespace ReportTitulList;

public static class Ext
{
    public static string Style { get; set; } = "";

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