using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            if (false)
                CreateTableWithVerticalMerge("testTable.docx");
            else
                CreateTableWithHorizontalMerge("testTable.docx");
        }

        public static void CreateTableWithVerticalMerge(string fileName)
        {
            // Use the file name and path passed in as an argument 
            // to open an existing Word 2007 document.

            using (WordprocessingDocument doc
                = WordprocessingDocument.Open(fileName, true))
            {
                Table table = new Table();
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 24
                        },
                        new BottomBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 24
                        },
                        new LeftBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 24
                        },
                        new RightBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 24
                        },
                        new InsideHorizontalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 24
                        },
                        new InsideVerticalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 24
                        }
                    )
                );
                table.AppendChild<TableProperties>(tblProp);


                TableRow tr = new TableRow();

                TableCell tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new TableCellProperties(new VerticalMerge() {Val = MergedCellValues.Restart }));
                tc.Append(new Paragraph(new Run(new Text("1"))));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new Paragraph(new Run(new Text("2"))));
                tr.Append(tc);

                table.Append(tr);


                tr = new TableRow();

                tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new TableCellProperties(new VerticalMerge() {Val = MergedCellValues.Continue }));
                tc.Append(new Paragraph());
                tr.Append(tc);

                // Create a second table cell by copying the OuterXml value of the first table cell.
                tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new Paragraph(new Run(new Text("22"))));
                tr.Append(tc);

                table.Append(tr);


                // Append the table to the document.
                doc.MainDocumentPart.Document.Body.Append(table);
            }
        }

        public static void CreateTableWithHorizontalMerge(string fileName)
        {
            using (WordprocessingDocument doc
                = WordprocessingDocument.Open(fileName, true))
            {
                Table table = new Table();
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 12
                        },
                        new BottomBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 12
                        },
                        new LeftBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 12
                        },
                        new RightBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 12
                        },
                        new InsideHorizontalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 12
                        },
                        new InsideVerticalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                            Size = 12
                        }
                    )
                );
                table.AppendChild<TableProperties>(tblProp);


                TableRow tr = new TableRow();

                TableCell tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new TableCellProperties(new GridSpan() {Val = 2 }));
                tc.Append(new Paragraph(new Run(new Text("1"))));
                tr.Append(tc);

                table.Append(tr);


                tr = new TableRow();

                tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new Paragraph(new Run(new Text("11"))));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto }));
                tc.Append(new Paragraph(new Run(new Text("22"))));
                tr.Append(tc);

                table.Append(tr);


                // Append the table to the document.
                doc.MainDocumentPart.Document.Body.Append(table);
            }
        }
    }
}
