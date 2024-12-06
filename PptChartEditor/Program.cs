using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace PptChartEditor;

public class Program
{
    public static void Main()
    {

        var chart1Data = new ChartData
        {
            SeriesNames = [
                                "Series 1",
                        "Series 2",
                        // "Series 3",
                        // "Series 4",
                        // "Series 5",
                        // "Series 6",
                        // "Series 7",
                        // "Series 8",
                        // "Series 9",
                        // "Series 10",
                        // "Series 11",
                    ],
            CategoryNames = [
                                "Category 1",
                        "Category 2",
                        "Category 3",
                        "Category 4",
                        "Category 5",
                        "Category 6",
                        "Category 7",
                        "Category 8",
                        "Category 9",
                        "Category 10",
                        "Category 11",
                    ],
            Values = [
                                [
                            1010, 1080, 1100, 1400, 1480, 1500, 1600, 1700, 1800, 1900, 2000
                        ],
                        [
                            10100, 10800, 11000, 14000, 14800, 15000, 16000, 17000, 18000, 19000, 20000
                        ],

                    ],
        };

        var name = "test";
        //var name = "acu2";

        var originalPptxPath = $"../../../../data/{name}-orig.pptx";
        var workingPptxPath = $"../../../../data/{name}.pptx";

        File.Copy(originalPptxPath, workingPptxPath, true);

        using (var ppt = PresentationDocument.Open(workingPptxPath, true))
        {

            ChartUpdater.UpdateChart(ppt, 1, 1, chart1Data);

        }
    }
}