using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;


var chart1Data = new ChartData
{
  SeriesNames = [
                        "1",
                        "2",
                    ],
  CategoryNames = [
                        "1",
                        "2",
                        "3",
                        "4",
                        "5",
                        "6",
                        "7",
                        "8",
                    ],
  Values = [
                        [
                            6000, 3100, 3200, 3300, 4530, 2300, 1200, 3450,
                        ],
                        [
                            2010, 2240, 2300, 2210, 2340, 1230, 3405, 2340,
                        ],
                        // [
                        //     180, 200, 220, 230, 234, 123, 345, 234,
                        // ],
                    ],
};

using (var ppt = PresentationDocument.Open("acu.pptx", true))
{

  ChartUpdater.UpdateChart(ppt, 9, 1, chart1Data);

}
