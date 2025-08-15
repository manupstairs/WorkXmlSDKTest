using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

namespace WorkXmlSDKTest
{
    public static class RadarChartProvider
    {
        public static Drawing CreateRadarChart(WordprocessingDocument doc)
        {
            var chartPart = doc.MainDocumentPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = "en-US" });

            var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
            chart.Append(new AutoTitleDeleted() { Val = true });

            var plotArea = new PlotArea();
            var radarChart = new RadarChart(
                new RadarStyle() { Val = RadarStyleValues.Marker }
            );

            var radarSeries = new RadarChartSeries(
                new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = 0U },
                new Order() { Val = 0U },
                new SeriesText(new NumericValue() { Text = "示例数据" })
            );

            // Category axis data (StringLiteral)
            var cat = new CategoryAxisData(
                new StringLiteral(
                    new PointCount() { Val = 3U },
                    new StringPoint() { Index = 0U, NumericValue = new NumericValue("类别A") },
                    new StringPoint() { Index = 1U, NumericValue = new NumericValue("类别B") },
                    new StringPoint() { Index = 2U, NumericValue = new NumericValue("类别C") }
                )
            );

            // Values (NumberLiteral)
            var val = new Values(
                new NumberLiteral(
                    new PointCount() { Val = 3U },
                    new NumericPoint() { Index = 0U, NumericValue = new NumericValue("10") },
                    new NumericPoint() { Index = 1U, NumericValue = new NumericValue("20") },
                    new NumericPoint() { Index = 2U, NumericValue = new NumericValue("30") }
                )
            );

            radarSeries.Append(cat);
            radarSeries.Append(val);
            radarChart.Append(radarSeries);

            // Axis IDs
            uint catAxisId = 48650112U;
            uint valAxisId = 48672768U;
            radarChart.Append(new AxisId() { Val = catAxisId });
            radarChart.Append(new AxisId() { Val = valAxisId });

            // Category Axis
            var categoryAxis = new CategoryAxis(
                new AxisId() { Val = catAxisId },
                new Scaling(new Orientation() { Val = OrientationValues.MinMax }),
                new AxisPosition() { Val = AxisPositionValues.Bottom },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = valAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new AutoLabeled() { Val = true },
                new LabelAlignment() { Val = LabelAlignmentValues.Center },
                new LabelOffset() { Val = 100 }
            );

            // Value Axis
            var valueAxis = new ValueAxis(
                new AxisId() { Val = valAxisId },
                new Scaling(new Orientation() { Val = OrientationValues.MinMax }),
                new AxisPosition() { Val = AxisPositionValues.Left },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = true },
                new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis() { Val = catAxisId },
                new Crosses() { Val = CrossesValues.AutoZero },
                new CrossBetween() { Val = CrossBetweenValues.Between }
            );

            plotArea.Append(radarChart);
            plotArea.Append(categoryAxis);
            plotArea.Append(valueAxis);
            chart.Append(plotArea);
            chartPart.ChartSpace.Append(chart);

            var drawing = new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5486400, Cy = 3200400 },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = 1U, Name = "Radar Chart" },
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

            return drawing;
        }
    }
}
