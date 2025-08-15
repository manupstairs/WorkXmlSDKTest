using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

class Program
{
    static void Main(string[] args)
    {
        CreateWordWithPieChart("WordWithPieChart.docx");
    }

    public static void CreateWordWithPieChart(string filePath)
    {
        using (var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            // Add chart part
            var chartPart = mainPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = "en-US" });

            var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
            chart.Append(new AutoTitleDeleted() { Val = true });

            var plotArea = new PlotArea();
            var pieChart = new PieChart();

            var pieSeries = new PieChartSeries(
                new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = 0U },
                new Order() { Val = 0U },
                new SeriesText(new NumericValue() { Text = "示例数据" })
            );

            // Category axis data
            var cat = new CategoryAxisData();
            var strRef = new StringReference();
            strRef.Append(new Formula() { Text = "" });
            strRef.Append(new StringCache(
                new PointCount() { Val = 3U },
                new StringPoint() { Index = 0U, NumericValue = new NumericValue("类别A") },
                new StringPoint() { Index = 1U, NumericValue = new NumericValue("类别B") },
                new StringPoint() { Index = 2U, NumericValue = new NumericValue("类别C") }
            ));
            cat.Append(strRef);

            // Values
            var val = new Values();
            var numRef = new NumberReference();
            numRef.Append(new Formula() { Text = "" });
            numRef.Append(new NumberingCache(
                new PointCount() { Val = 3U },
                new NumericPoint() { Index = 0U, NumericValue = new NumericValue("30") },
                new NumericPoint() { Index = 1U, NumericValue = new NumericValue("50") },
                new NumericPoint() { Index = 2U, NumericValue = new NumericValue("20") }
            ));
            val.Append(numRef);

            pieSeries.Append(cat);
            pieSeries.Append(val);
            pieChart.Append(pieSeries);
            plotArea.Append(pieChart);
            chart.Append(plotArea);
            chartPart.ChartSpace.Append(chart);

            // Add chart to document
            var drawing = new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5486400, Cy = 3200400 },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                    {
                        Id = 1U,
                        Name = "Pie Chart"
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                        new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }
                    ),
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new ChartReference() { Id = doc.MainDocumentPart.GetIdOfPart(chartPart) }
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                )
            );

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(drawing)));

            mainPart.Document.Save();
        }
    }
}
