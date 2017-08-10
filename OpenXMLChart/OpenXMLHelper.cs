using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLChart
{
    public class OpenXMLHelper
    {
        /// <summary>
        /// draw the 2D bar chart
        /// index start from 1
        /// </summary>
        /// <param name="startx">index start from 1 for row</param>
        /// <param name="starty">index start from 1 for column</param>
        /// <param name="columnCount"></param>
        /// <param name="rowCount"></param>
        public void InsertChartInSpreadsheet(WorksheetPart sheetpart,string sheetName,int startx, int starty, int columnCount, int rowCount,int chart_pointx,int chart_pointy)
        {
            WorksheetPart worksheetPart = CurrentWorksheetPart;
            #region SDK How to example code
            // Add a new drawing to the worksheet.
            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            worksheetPart.Worksheet.Save();
            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());
            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));
            #endregion
            string sheetName = GetCurrentSheetName();
            string columnName = GetColumnName(starty - 1);
            string formulaString = string.Format("{0}!${1}${2}:${3}${4}", sheetName, columnName, startx + 1, columnName, startx + rowCount - 1);
            CategoryAxisData cad = new CategoryAxisData();
            cad.StringReference = new StringReference() { Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula(formulaString) };
            uint i = 0;
            for (int sIndex = 1; sIndex < columnCount; sIndex++)
            {
                columnName = GetColumnName(starty + sIndex - 1);
                formulaString = string.Format("{0}!${1}${2}", sheetName, columnName, startx);
                SeriesText st = new SeriesText();
                st.StringReference = new StringReference() { Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula(formulaString) };
                formulaString = string.Format("{0}!${1}${2}:${3}${4}", sheetName, columnName, startx + 1, columnName, startx + rowCount - 1);
                DocumentFormat.OpenXml.Drawing.Charts.Values v = new DocumentFormat.OpenXml.Drawing.Charts.Values();
                v.NumberReference = new NumberReference() { Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula(formulaString) };
                BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index() { Val = new UInt32Value(i) },
                    new Order() { Val = new UInt32Value(i) }, st, v));
                if (sIndex == 1)
                    barChartSeries.AppendChild(cad);
                i++;
            }
            #region SDK how to  example Code
            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });
            // Add the Category Axis.
            CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId() { Val = new UInt32Value(48650112u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                        DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = new UInt32Value(48672768U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new AutoLabeled() { Val = new BooleanValue(true) },
                new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                new LabelOffset() { Val = new UInt16Value((ushort)100) }));
            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                        DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberFormat() { FormatCode = new StringValue("General"), SourceLinked = new BooleanValue(true) },
                new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));
            // Add the chart Legend.
            Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
                new Layout()));
            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });
            // Save the chart part.
            chartPart.ChartSpace.Save();
            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            twoCellAnchor.Append(new FromMarker(new ColumnId("9"),
                new ColumnOffset("581025"),
                new RowId("17"),
                new RowOffset("114300")));
            twoCellAnchor.Append(new ToMarker(new ColumnId("17"),
                new ColumnOffset("276225"),
                new RowId("32"),
                new RowOffset("0")));
            // Append a GraphicFrame to the TwoCellAnchor object.
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";
            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));
            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                    new Extents() { Cx = 0L, Cy = 0L }));
            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));
            twoCellAnchor.Append(new ClientData());
            // Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save();
            #endregion
        }
        public static string GetColumnName(int index)
        {
            if (index < 0) { throw new Exception("invalid parameter"); }
            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);
            return String.Join(string.Empty, chars.ToArray());
        }
    }
}
