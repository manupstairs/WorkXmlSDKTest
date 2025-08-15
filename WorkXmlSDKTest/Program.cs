using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using WorkXmlSDKTest;

class Program
{
    static void Main(string[] args)
    {
        CreateWordWithPieChart("WordWithCharts.docx");
    }

    public static void CreateWordWithPieChart(string filePath)
    {
        using (var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            // 插入曲线图
            var lineDrawing = LineChartProvider.CreateLineChart(doc);
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(lineDrawing)));

            var barDrawing = BarChartProvider.CreateBarChart(doc);
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(barDrawing)));

            var pieDrawing = PieChartProvider.CreatePieChart(doc);
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(pieDrawing)));

            var radarDrawing = RadarChartProvider.CreateRadarChart(doc);
            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(radarDrawing)));

            mainPart.Document.Save();
        }
    }
}
