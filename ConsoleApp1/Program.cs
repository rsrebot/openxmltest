using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Spire.Presentation;
using Spire.Presentation.Charts;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;


using (var ppt = PresentationDocument.Open("chart.pptx", true))
{

  var t = ppt.PresentationPart.SlideParts.Where(p => p.ChartParts.Any()).First();

  var c = t.ChartParts.First();

  using (SpreadsheetDocument ssDoc = SpreadsheetDocument.Open(c.EmbeddedPackagePart.GetStream(FileMode.Open, FileAccess.ReadWrite), true))
  {
    var wbp = ssDoc.WorkbookPart;
    Sheet sheet = wbp.Workbook.Descendants<Sheet>().First();
    Worksheet ws = ((WorksheetPart)wbp.GetPartById(sheet.Id)).Worksheet;
    var cell = ws.GetFirstChild<SheetData>().Elements<Row>().First().Elements<Cell>().First();
    cell.CellValue = new CellValue("10000");

    ssDoc.Save();
  }
}


//ppt.Save();


// Presentation ppt1 = new();

// ppt1.LoadFromFile("chart.pptx");

// foreach (Shape s in ppt1.Slides[0].Shapes)
// {
//   var c = s as IChart;
//   if (c != null)
//   {
//     Console.WriteLine(c.ChartData);
//     Console.WriteLine(c.Type);
//     Console.Write(c.Series[0].Values[0].Value);
//     c.Series[0].Values[0].Value = 10000;
//   }
// }
//ppt1.SaveToFile("chart.pptx", FileFormat.Pptx2013);

//"/ppt/charts/chart1.xml"
//"/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"