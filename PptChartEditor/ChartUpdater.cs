using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Presentation;

public class ChartData
{
  public required string[] SeriesNames;
  public required string[] CategoryNames;
  public required double[][] Values;
}


public class ChartUpdater
{
  public static void UpdateChart(PresentationDocument pDoc, int slideNumber, int chartIdx, ChartData chartData)
  {
    if (pDoc.PresentationPart == null)
    {
      throw new ArgumentException("Invalid document");
    }

    PresentationPart mainDocumentPart = pDoc.PresentationPart;

    if (mainDocumentPart.Presentation?.SlideIdList?.Skip(slideNumber - 1).First() is not SlideId slideId)
    {
      throw new ArgumentException("Invalid slide number");
    }

    if (slideId.RelationshipId is not StringValue relationshipId || relationshipId is null || relationshipId.InnerText is null)
    {
      throw new ArgumentException("Invalid slide number");
    }

    SlidePart? slide = mainDocumentPart?.GetPartById(relationshipId.InnerText) as SlidePart;

    if (slide == null)
    {
      throw new ArgumentException("Invalid slide number");
    }

    var chart = slide.ChartParts.Skip(chartIdx - 1).First();

    if (chart != null)
    {
      UpdateChart(chart, chartData);
    }
  }

  private static void UpdateChart(ChartPart chartPart, ChartData chartData)
  {
    if (chartData.Values.Length != chartData.SeriesNames.Length)
      throw new ArgumentException("Invalid chart data");
    foreach (var ser in chartData.Values)
    {
      if (ser.Length != chartData.CategoryNames.Length)
        throw new ArgumentException("Invalid chart data");
    }

    UnlinkSpreadsheet(chartPart);
    UpdateSeries(chartPart, chartData);
  }

  private static void UnlinkSpreadsheet(ChartPart chartPart)
  {
    XDocument cpx = chartPart.GetXDocument();
    XElement newRoot = (XElement)UnlinkSpreadsheetTransform(cpx.Root);
    cpx.Root.ReplaceWith(newRoot);
    //chartPart.PutXDocument();
    chartPart.SaveXDocument();
    var workbookPartPair = chartPart.Parts.FirstOrDefault(p => p.OpenXmlPart.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    if (workbookPartPair != null)
      chartPart.DeletePart(workbookPartPair.OpenXmlPart);
  }

  private static void UpdateSeries(ChartPart chartPart, ChartData chartData)
  {
    XDocument cpXDoc = chartPart.GetXDocument();
    XElement root = cpXDoc.Root;
    var firstSeries = root.Descendants(C.ser).FirstOrDefault();
    var numLit = firstSeries.Elements(C.val).Elements(C.numLit).FirstOrDefault();

    // remove all but first series
    firstSeries.Parent.Elements(C.ser).Skip(1).Remove();

    var newSetOfSeries = chartData.SeriesNames
        .Select((string sn, int si) =>
        {
          XElement cat = null;

          if (firstSeries.Elements(C.cat).Elements(C.numLit).Any())
          {
            cat = new XElement(C.cat,
                    new XElement(C.numLit,
                        firstSeries.Elements(C.cat).Elements(C.numLit).Elements(C.formatCode),
                        new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                        chartData.CategoryNames.Select((string cn, int ci) =>
                        {
                          var newPt = new XElement(C.pt,
                                  new XAttribute("idx", ci),
                                  new XElement(C.v,
                                      Int32.Parse(chartData.CategoryNames[ci]).ToString()));  // convert to int and back to string
                                                                                              // to make sure that the cat names are integer values, i.e. dates
                          return newPt;
                        })));
          }
          else
          {
            cat = new XElement(C.cat,
                    new XElement(C.strLit,
                        new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                        chartData.CategoryNames.Select((string cn, int ci) =>
                        {
                          var newPt = new XElement(C.pt,
                                  new XAttribute("idx", ci),
                                  new XElement(C.v, chartData.CategoryNames[ci]));
                          return newPt;
                        })));
          }

          var newSer = new XElement(C.ser,
                  new XElement(C.idx, new XAttribute("val", si)),
                  new XElement(C.order, new XAttribute("val", si)),
                  new XElement(C.tx,
                      new XElement(C.v, sn)),
                  firstSeries.Elements().Where(e => e.Name == C.spPr),
                  cat,
                  new XElement(C.val,
                      new XElement(C.numLit,
                          numLit.Elements(C.formatCode),
                          new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                          chartData.CategoryNames.Select((string cn, int ci) =>
                          {
                            var newPt = new XElement(C.pt,
                                    new XAttribute("idx", ci),
                                    new XElement(C.v, chartData.Values[si][ci]));
                            return newPt;
                          }))),
                  firstSeries.Elements().Where(e => e.Name != C.idx &&
                      e.Name != C.order &&
                      e.Name != C.tx &&
                      e.Name != C.cat &&
                      e.Name != C.val &&
                      e.Name != C.spPr)
              );
          int accentNumber = (si % 6) + 1;
          newSer = (XElement)UpdateAccentTransform(newSer, accentNumber);
          return newSer;
        });
    firstSeries.ReplaceWith(newSetOfSeries);
    //chartPart.PutXDocument();
    chartPart.SaveXDocument();
  }

  private static object UpdateAccentTransform(XNode node, int accentNumber)
  {
    XElement element = node as XElement;
    if (element != null)
    {
      if (element.Name == A.schemeClr && (string)element.Attribute("val") == "accent1")
        return new XElement(A.schemeClr, new XAttribute("val", "accent" + accentNumber));

      return new XElement(element.Name,
          element.Attributes(),
          element.Nodes().Select(n => UpdateAccentTransform(n, accentNumber)));
    }
    return node;
  }

  private static object UnlinkSpreadsheetTransform(XNode node)
  {
    XElement element = node as XElement;
    if (element != null)
    {
      if (element.Name == C.externalData)
        return null;

      if (element.Name == C.numFmt)
        return new XElement(element.Name,
            element.Attribute("formatCode"));

      if (element.Name == C.numRef)
        return new XElement(C.numLit,
            element.Elements(C.numCache).Elements());

      if (element.Name == C.strRef)
        return new XElement(C.strLit,
            element.Elements(C.strCache).Elements());

      if (element.Name == C.tx && element.Parent.Name == C.ser)
        return new XElement(C.tx, element.Descendants(C.v));

      return new XElement(element.Name,
          element.Attributes(),
          element.Nodes().Select(n => UnlinkSpreadsheetTransform(n)));
    }
    return node;
  }

  internal static bool UpdateChart(PresentationDocument pDoc, int slideNumber, ChartData chartData)
  {
    var presentationPart = pDoc.PresentationPart;
    var pXDoc = presentationPart.GetXDocument();
    var sldIdElement = pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId).Skip(slideNumber - 1).FirstOrDefault();
    if (sldIdElement != null)
    {
      var rId = (string)sldIdElement.Attribute(R.id);
      var slidePart = presentationPart.GetPartById(rId);
      var sXDoc = slidePart.GetXDocument();
      var chartRid = (string)sXDoc.Descendants(C.chart).Attributes(R.id).FirstOrDefault();
      if (chartRid != null)
      {
        ChartPart chartPart = (ChartPart)slidePart.GetPartById(chartRid);
        UpdateChart(chartPart, chartData);
        return true;
      }
      return true;
    }
    return false;
  }
}