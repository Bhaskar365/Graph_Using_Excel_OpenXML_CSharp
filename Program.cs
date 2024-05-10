// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

Console.WriteLine("Hello, World!");

CreateExcelDoc();
void CreateExcelDoc(string fileName = "C:\\Graph_Using_Excel_Openxml_C#\\ConsoleApp1\\ConsoleApp1\\NewFolder\\sample2.xlsx")
{
    List<Student> students = new List<Student>();
    Initizalize(students);

    using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
    {
        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet();

        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Students" };

        SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

        // Add drawing part to WorksheetPart
        DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        worksheetPart.Worksheet.Save();

        drawingsPart.WorksheetDrawing = new WorksheetDrawing();

        sheets.Append(sheet);

        workbookPart.Workbook.Save();

        // Add a new chart and set the chart language
        ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = new ChartSpace();
        chartPart.ChartSpace.AppendChild(new EditingLanguage() { Val = "en-US" });
        Chart chart = chartPart.ChartSpace.AppendChild(new Chart());
        chart.AppendChild(new AutoTitleDeleted() { Val = true }); // We don't want to show the chart title

        // Create a new Clustered Column Chart
        PlotArea plotArea = chart.AppendChild(new PlotArea());
        Layout layout = plotArea.AppendChild(new Layout());

        BarChart barChart = plotArea.AppendChild(new BarChart(
                new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) },
                new VaryColors() { Val = false }
            ));

        // Constructing header
        Row row = new Row();
        int rowIndex = 1;

        row.AppendChild(ConstructCell(string.Empty, CellValues.String));

        foreach (var month in Months.Short)
        {
            row.AppendChild(ConstructCell(month, CellValues.String));
        }

        // Insert the header row to the Sheet Data
        sheetData.AppendChild(row);
        rowIndex++;

        // Create chart series
        for (int i = 0; i < students.Count; i++)
        {
            BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(
                new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (uint)i },
                new Order() { Val = (uint)i },
                new SeriesText(new NumericValue() { Text = students[i].Name })
            ));

            // Adding category axis to the chart
            CategoryAxisData categoryAxisData = barChartSeries.AppendChild(new CategoryAxisData());

            // Category
            // Constructing the chart category
            string formulaCat = "Students!$B$1:$G$1";

            StringReference stringReference = categoryAxisData.AppendChild(new StringReference()
            {
                Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaCat }
            });

            StringCache stringCache = stringReference.AppendChild(new StringCache());
            stringCache.Append(new PointCount() { Val = (uint)Months.Short.Length });

            for (int j = 0; j < Months.Short.Length; j++)
            {
                stringCache.AppendChild(new NumericPoint() { Index = (uint)j }).Append(new NumericValue(Months.Short[j]));
            }
        }

        var chartSeries = barChart.Elements<BarChartSeries>().GetEnumerator();

        for (int i = 0; i < students.Count; i++)
        {
            row = new Row();

            row.AppendChild(ConstructCell(students[i].Name, CellValues.String));

            chartSeries.MoveNext();

            string formulaVal = string.Format("Students!$B${0}:$D${0}", rowIndex);
            DocumentFormat.OpenXml.Drawing.Charts.Values values = chartSeries.Current.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());

            NumberReference numberReference = values.AppendChild(new NumberReference()
            {
                Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formulaVal }
            });

            NumberingCache numberingCache = numberReference.AppendChild(new NumberingCache());
            numberingCache.Append(new PointCount() { Val = (uint)Months.Short.Length });

            for (uint j = 0; j < students[i].Values.Length; j++)
            {
                var value = students[i].Values[j];

                row.AppendChild(ConstructCell(value.ToString(), CellValues.Number));

                numberingCache.AppendChild(new NumericPoint() { Index = j }).Append(new NumericValue(value.ToString()));
            }

            sheetData.AppendChild(row);
            rowIndex++;
        }

        barChart.AppendChild(new DataLabels(
                            new ShowLegendKey() { Val = false },
                            new ShowValue() { Val = false },
                            new ShowCategoryName() { Val = false },
                            new ShowSeriesName() { Val = false },
                            new ShowPercent() { Val = false },
                            new ShowBubbleSize() { Val = false }
                        ));

        barChart.Append(new AxisId() { Val = 48650112u });
        barChart.Append(new AxisId() { Val = 48672768u });

        // Adding Category Axis
        plotArea.AppendChild(
            new CategoryAxis(
                new AxisId() { Val = 48650112u },
                new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                new Delete() { Val = false },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = 48672768u },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new AutoLabeled() { Val = true },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) }
            ));

        // Adding Value Axis
        plotArea.AppendChild(
            new ValueAxis(
                new AxisId() { Val = 48672768u },
                new Scaling(new Orientation() { Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax) }),
                new Delete() { Val = false },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                {
                    FormatCode = "General",
                    SourceLinked = true
                },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = 48650112u },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }
            ));

        chart.Append(
                new PlotVisibleOnly() { Val = true },
                new DisplayBlanksAs() { Val = new EnumValue<DisplayBlanksAsValues>(DisplayBlanksAsValues.Gap) },
                new ShowDataLabelsOverMaximum() { Val = false }
            );

        chartPart.ChartSpace.Save();

        // Positioning the chart on the spreadsheet
        TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());

        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                new ColumnId("5"),
                new ColumnOffset("0"),
                new RowId((rowIndex + 2).ToString()),
                new RowOffset("10")
            ));

        twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                new ColumnId("12"),
                new ColumnOffset("0"),
                new RowId((rowIndex + 12).ToString()),
                new RowOffset("40")
            ));

        // Append GraphicFrame to TwoCellAnchor
        GraphicFrame graphicFrame = twoCellAnchor.AppendChild(new GraphicFrame());
        graphicFrame.Macro = string.Empty;

        graphicFrame.Append(new NonVisualGraphicFrameProperties(
                new NonVisualDrawingProperties()
                {
                    Id = 2u,
                    Name = "Sample Chart"
                },
                new NonVisualGraphicFrameDrawingProperties()
            ));

        graphicFrame.Append(new Transform(
                new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                new DocumentFormat.OpenXml.Drawing.Extents() { Cx = 0L, Cy = 0L }
            ));

        graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
                new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }
                    )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            ));

        twoCellAnchor.Append(new ClientData());

        drawingsPart.WorksheetDrawing.Save();

        worksheetPart.Worksheet.Save();
    }
}

Cell ConstructCell(string value, CellValues dataType)
{
    return new Cell()
    {
        CellValue = new CellValue(value),
        DataType = new EnumValue<CellValues>(dataType),
    };
}

void Initizalize(List<Student> students)
{
    students.AddRange(new Student[] {
                new Student
                {
                    Name = "Liza",
                    Values = new byte[] { 10, 25, 30 }
                },
                new Student
                {
                    Name = "Macy",
                    Values = new byte[] { 20, 15, 26 }
                },
                new Student
                {
                    Name = "Razer",
                    Values = new byte[] { 20, 45, 32 }
                }
            });
}

class Student
{
    public string Name { get; set; }
    public byte[] Values { get; set; }
}

struct Months
{
    public static string[] Short = {
            "Label 1",
            "Label 2",
            "Label 3",
        };
}
