using QuestPDF.Infrastructure;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

QuestPDF.Settings.License = LicenseType.Community;

string bladeSelection = "Foil";
int num_pools = 3;
List<string> squads = new List<string> {"UT Austin A", "UT Austin C", "UT Dallas A", "Texas A&M A", "TXST B", "TXST C", "UNT", "OU"};
int num_squads = squads.Count;

QuestPDF.Fluent.Document
    .Create(container =>
    {
        container.Page(page =>
        {
            page.Size(PageSizes.Letter);
            page.Margin(0.5f, Unit.Inch);

            page.Header()
                .Text($"{bladeSelection} Pool #{num_pools} Summary")
                .FontSize(28)
                .Bold();

            page.Content().Column(column =>
            {
                column.Item().Height(20);
                column.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.ConstantColumn(20);
                        columns.ConstantColumn(120);
                        for (int i = 0; i < num_squads; i++)
                        {
                            columns.RelativeColumn(1);
                        }
                    });

                    table.Cell().Element(CellStyle).Text("#").Bold().AlignCenter();
                    table.Cell().Element(CellStyle).Text("Squad").Bold().AlignCenter();

                    for (int i = 0; i < num_squads; i++)
                        table.Cell().Element(CellStyle).Text($"{i + 1}").Bold().AlignCenter();

                    for (int i = 0; i < num_squads; i++)
                    {
                        table.Cell().Element(CellStyle).Text($"{i + 1}").Bold().AlignCenter();
                        table.Cell().Element(CellStyle).Text(squads[i]).AlignCenter();

                        for (int j = 0; j < num_squads; j++)
                        {
                            if (i == j)
                                table.Cell().Border(1).Background(Colors.Grey.Darken1).AlignCenter();
                            else
                                table.Cell().Element(CellStyle).Text(" ").AlignCenter();
                        }
                    }

                    static IContainer CellStyle(IContainer container)
                        => container.Border(1).Padding(8);
                });
                column.Item().Height(20);
                column.Item().Row(row =>
                {
                    row.RelativeItem();
                    row.ConstantItem(522).Image("./poolBoutSequence.png");
                    row.RelativeItem();
                });
            });
        });
    }).GeneratePdf("./outputPdf/1poolSummary.pdf");


QuestPDF.Fluent.Document
    .Create(container => 
    {
        container.Page(page =>
        {
            page.Size(PageSizes.Letter);
            page.Margin(0.2f, Unit.Inch);

            page.Content().Column(column =>
            {
                column.Item().Row(row =>
                {
                    row.RelativeItem()
                        .Text("Encounter Summary")
                        .FontSize(15)
                        .Bold()
                        .AlignLeft();

                    row.RelativeItem()
                        .Text("[blade] Pool #[poolNum]")
                        .FontSize(15)
                        .Bold()
                        .AlignRight();
                });

                column.Item().Row(row =>
                {
                    row.RelativeItem()
                        .Text("Squad1 vs. Squad2")
                        .FontSize(15)
                        .AlignLeft();

                    row.RelativeItem()
                        .Text("Seed1 vs. Seed2")
                        .FontSize(15)
                        .AlignRight();
                });

                column.Item().Height(20);

                column.Item().Row(row =>
                {
                    row.RelativeItem();

                row.ConstantItem(560).Table(table =>
                {
                    static IContainer CellStyle_Header(IContainer container)
                            => container
                                .BorderBottom(2)
                                .AlignBottom()
                                .AlignCenter();
                    static IContainer NoTopBottomBorder(IContainer container)
                        => container
                            .BorderLeft(1)
                            .BorderRight(1)
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    static IContainer NoTopBorder(IContainer container)
                        => container
                            .BorderLeft(1)
                            .BorderRight(1)
                            .BorderBottom(1)
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    static IContainer NoBorder(IContainer container)
                        => container
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    static IContainer FullBorder(IContainer container)
                        => container
                            .Border(1)
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    table.ColumnsDefinition(columns =>
                    {
                        columns.ConstantColumn(17);
                        columns.ConstantColumn(96);
                        columns.ConstantColumn(86);
                        columns.ConstantColumn(17);
                        
                        for (int i = 0; i < 4; i++)
                            columns.ConstantColumn(32);
                        
                        columns.ConstantColumn(17);
                        columns.ConstantColumn(86);
                        columns.ConstantColumn(96);
                        columns.ConstantColumn(17);
                    });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyle_Header).Text("#").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Squad Left").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Fencers").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("SI.").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Score").Bold().FontSize(12).AlignCenter();

                            header.Cell().ColumnSpan(2).Element(CellStyle_Header)
                                .Text("Team Victories").Bold().FontSize(12).AlignCenter();

                            header.Cell().Element(CellStyle_Header).Text("Score").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("SI").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Fencers").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Squad Right").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("#").Bold().FontSize(12).AlignCenter();
                        });
                        // first three row
                        List<string> strips = new List<string> {"C", "B", "A"};
                        for (int i = 0; i < 3; i++){
                            if (i == 2) {
                                table.Cell().Element(NoTopBottomBorder).Text("1");
                                table.Cell().Element(NoTopBottomBorder).Text("UT Austin A");
                            } else {
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                            }
                            table.Cell().Element(FullBorder).Text("");
                            table.Cell().Element(FullBorder).Text($"{strips[i]}");
                            for (int j = 0; j < 4; j++) {
                                table.Cell().Element(FullBorder).Text(" ");
                            }
                            table.Cell().Element(FullBorder).Text($"{strips[i]}");
                            table.Cell().Element(FullBorder).Text(" ");
                            if (i == 2) {
                                table.Cell().Element(NoTopBottomBorder).Text("UT Austin C");
                                table.Cell().Element(NoTopBottomBorder).Text("2");
                            } else {
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                            }
                        }

                        // fourth row
                        table.Cell().Element(NoTopBottomBorder).Text(" ");
                        table.Cell().Element(NoTopBottomBorder).Text(" ");
                        for (int i = 0; i < 8; i++) {
                            table.Cell().Element(NoBorder).Text(" ");
                        }
                        table.Cell().Element(NoTopBottomBorder).Text(" ");
                        table.Cell().Element(NoTopBottomBorder).Text(" ");

                        // fifth row
                        table.Cell().Element(NoTopBorder).Text(" ");
                        table.Cell().Element(NoTopBorder).Text(" ");
                        table.Cell().Element(FullBorder).Text(" ");
                        table.Cell().Element(FullBorder).Text("Alt");
                        for (int i = 0; i < 4; i++) {
                            table.Cell().Element(NoBorder).Text(" ");
                        }
                        table.Cell().Element(FullBorder).Text("Alt");
                        table.Cell().Element(FullBorder).Text(" ");
                        table.Cell().Element(NoTopBorder).Text(" ");
                        table.Cell().Element(NoTopBorder).Text(" ");
                    });
                    row.RelativeItem();
                });
                column.Item().Height(20);
                column.Item().Image("./signatureFields.png");
                column.Item().Height(20);
                column.Item().LineHorizontal(2).LineColor(Colors.Black);
                
                column.Item().Height(20);
                column.Item().Row(row =>
                {
                    row.RelativeItem()
                        .Text("Encounter Summary")
                        .FontSize(15)
                        .Bold()
                        .AlignLeft();

                    row.RelativeItem()
                        .Text("[blade] Pool #[poolNum]")
                        .FontSize(15)
                        .Bold()
                        .AlignRight();
                });

                column.Item().Row(row =>
                {
                    row.RelativeItem()
                        .Text("Squad3 vs. Squad4")
                        .FontSize(15)
                        .AlignLeft();

                    row.RelativeItem()
                        .Text("Seed3 vs. Seed4")
                        .FontSize(15)
                        .AlignRight();
                });

                column.Item().Height(20);
                column.Item().Row(row =>
                {
                    row.RelativeItem();

                row.ConstantItem(560).Table(table =>
                {
                    static IContainer CellStyle_Header(IContainer container)
                            => container
                                .BorderBottom(2)
                                .AlignBottom()
                                .AlignCenter();
                    static IContainer NoTopBottomBorder(IContainer container)
                        => container
                            .BorderLeft(1)
                            .BorderRight(1)
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    static IContainer NoTopBorder(IContainer container)
                        => container
                            .BorderLeft(1)
                            .BorderRight(1)
                            .BorderBottom(1)
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    static IContainer NoBorder(IContainer container)
                        => container
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    static IContainer FullBorder(IContainer container)
                        => container
                            .Border(1)
                            .AlignCenter()
                            .AlignMiddle()
                            .Height(30);
                    table.ColumnsDefinition(columns =>
                    {
                        columns.ConstantColumn(17);
                        columns.ConstantColumn(96);
                        columns.ConstantColumn(86);
                        columns.ConstantColumn(17);
                        
                        for (int i = 0; i < 4; i++)
                            columns.ConstantColumn(32);
                        
                        columns.ConstantColumn(17);
                        columns.ConstantColumn(86);
                        columns.ConstantColumn(96);
                        columns.ConstantColumn(17);
                    });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyle_Header).Text("#").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Squad Left").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Fencers").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("SI.").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Score").Bold().FontSize(12).AlignCenter();

                            header.Cell().ColumnSpan(2).Element(CellStyle_Header)
                                .Text("Team Victories").Bold().FontSize(12).AlignCenter();

                            header.Cell().Element(CellStyle_Header).Text("Score").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("SI").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Fencers").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("Squad Right").Bold().FontSize(12).AlignCenter();
                            header.Cell().Element(CellStyle_Header).Text("#").Bold().FontSize(12).AlignCenter();
                        });
                        // first three row
                        List<string> strips = new List<string> {"C", "B", "A"};
                        for (int i = 0; i < 3; i++){
                            if (i == 2) {
                                table.Cell().Element(NoTopBottomBorder).Text("3");
                                table.Cell().Element(NoTopBottomBorder).Text("UT Austin B");
                            } else {
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                            }
                            table.Cell().Element(FullBorder).Text("Dayou Ren");
                            table.Cell().Element(FullBorder).Text($"{strips[i]}");
                            for (int j = 0; j < 4; j++) {
                                table.Cell().Element(FullBorder).Text(" ");
                            }
                            table.Cell().Element(FullBorder).Text($"{strips[i]}");
                            table.Cell().Element(FullBorder).Text(" ");
                            if (i == 2) {
                                table.Cell().Element(NoTopBottomBorder).Text("UT Austin D");
                                table.Cell().Element(NoTopBottomBorder).Text("4");
                            } else {
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                                table.Cell().Element(NoTopBottomBorder).Text(" ");
                            }
                        }

                        // fourth row
                        table.Cell().Element(NoTopBottomBorder).Text(" ");
                        table.Cell().Element(NoTopBottomBorder).Text(" ");
                        for (int i = 0; i < 8; i++) {
                            table.Cell().Element(NoBorder).Text(" ");
                        }
                        table.Cell().Element(NoTopBottomBorder).Text(" ");
                        table.Cell().Element(NoTopBottomBorder).Text(" ");

                        // fifth row
                        table.Cell().Element(NoTopBorder).Text(" ");
                        table.Cell().Element(NoTopBorder).Text(" ");
                        table.Cell().Element(FullBorder).Text(" ");
                        table.Cell().Element(FullBorder).Text("Alt");
                        for (int i = 0; i < 4; i++) {
                            table.Cell().Element(NoBorder).Text(" ");
                        }
                        table.Cell().Element(FullBorder).Text("Alt");
                        table.Cell().Element(FullBorder).Text(" ");
                        table.Cell().Element(NoTopBorder).Text(" ");
                        table.Cell().Element(NoTopBorder).Text(" ");
                    });
                    row.RelativeItem();
                });
                column.Item().Height(20);
                column.Item().Image("./signatureFields.png");
            });
        });
    }).GeneratePdf("./outputPdf/2boutSheets.pdf");

string inputFolder = "./outputPdf";
string outputFileName = "combined.pdf";
using (PdfDocument outputDoc = new PdfDocument()) {
    foreach (string file in Directory.GetFiles(inputFolder, "*.pdf")) {
        using (PdfDocument inputDoc = PdfReader.Open(file, PdfDocumentOpenMode.Import)) {
            for (int i = 0; i < inputDoc.PageCount; i++) {
                outputDoc.AddPage(inputDoc.Pages[i]);
            }
        }
    }
    outputDoc.Save(Path.Combine(inputFolder, outputFileName));
}

foreach (string file in Directory.GetFiles(inputFolder, "*.pdf"))
{
    if (!Path.GetFileName(file).Equals(outputFileName, System.StringComparison.OrdinalIgnoreCase))
    {
        File.Delete(file);
    }
}