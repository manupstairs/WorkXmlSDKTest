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
        CreateWordWithPieChart("WordWithPieChart.docx");
    }

    public static void CreateWordWithPieChart(string filePath)
    {
        using (var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            var drawing = PieChartProvider.CreatePieChart(doc);

            body.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(drawing)));

            mainPart.Document.Save();
        }
    }
}
