using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static WordprocessingDocument CreateBarChart(WordprocessingDocument document)//List<ChartSubArea> chartList,
        {
            string title = "New Chart";

            Dictionary<string, int> data = new Dictionary<string, int>();
            data.Add("abc", 1);

            // Get MainDocumentPart of Document
            MainDocumentPart mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Create ChartPart object in Word Document
            ChartPart chartPart = mainPart.AddNewPart<ChartPart>("rId110");

            // the root element of chartPart 
            dc.ChartSpace chartSpace = new dc.ChartSpace();
            chartSpace.Append(new dc.EditingLanguage() { Val = "en-us" });

            // Create Chart 
            dc.Chart chart = new dc.Chart();
            chart.Append(new dc.AutoTitleDeleted() { Val = true });

            // Define the 3D view
            dc.View3D view3D = new dc.View3D();
            view3D.Append(new dc.RotateX() { Val = 30 });
            view3D.Append(new dc.RotateY() { Val = 0 });

            // Intiliazes a new instance of the PlotArea class
            dc.PlotArea plotArea = new dc.PlotArea();
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection()
            { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
               new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

            plotArea.Append(new dc.Layout());


            dc.ChartShapeProperties chartShapePros = new dc.ChartShapeProperties();

            uint i = 0;
            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.
            foreach (string key in data.Keys)
            {
                BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
                {
                    Val =
                    new UInt32Value(i)
                },
                    new Order() { Val = new UInt32Value(i) },
                    new SeriesText(new NumericValue() { Text = key })));

                StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
                strLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(0U) }).Append(new NumericValue(title));

                NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
                numLit.Append(new FormatCode("General"));
                numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
                numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append
                (new NumericValue(data[key].ToString()));

                i++;
            }

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis.
            CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
            { Val = new UInt32Value(48650112u) }, new Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation()
            {
                Val = new EnumValue<DocumentFormat.
                OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
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
            new Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
            new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
            new MajorGridlines(),
            new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
            {
                FormatCode = new StringValue("General"),
                SourceLinked = new BooleanValue(true)
            }, new TickLabelPosition()
            {
                Val = new EnumValue<TickLabelPositionValues>
                (TickLabelPositionValues.NextTo)
            }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
            new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            // create child elements of the c:legend element
            dc.Legend legend = new dc.Legend();
            legend.Append(new dc.LegendPosition() { Val = LegendPositionValues.Right });
            dc.Overlay overlay = new dc.Overlay() { Val = false };
            legend.Append(overlay);

            dc.TextProperties textPros = new DocumentFormat.OpenXml.Drawing.Charts.TextProperties();
            textPros.Append(new d.BodyProperties());
            textPros.Append(new d.ListStyle());

            d.Paragraph paragraph = new d.Paragraph();
            d.ParagraphProperties paraPros = new d.ParagraphProperties();
            d.DefaultParagraphProperties defaultParaPros = new d.DefaultParagraphProperties();
            defaultParaPros.Append(new d.LatinFont() { Typeface = "Arial", PitchFamily = 34, CharacterSet = 0 });
            defaultParaPros.Append(new d.ComplexScriptFont() { Typeface = "Arial", PitchFamily = 34, CharacterSet = 0 });
            paraPros.Append(defaultParaPros);
            paragraph.Append(paraPros);
            paragraph.Append(new d.EndParagraphRunProperties() { Language = "en-Us" });

            textPros.Append(paragraph);
            legend.Append(textPros);

            // Append c:view3D, c:plotArea and c:legend elements to the end of c:chart element
            chart.Append(view3D);
            chart.Append(plotArea);
            chart.Append(legend);

            // Append the c:chart element to the end of c:chartSpace element
            chartSpace.Append(chart);

            // Create c:spPr Elements and fill the child elements of it
            chartShapePros = new dc.ChartShapeProperties();
            d.Outline outline = new d.Outline();
            outline.Append(new d.NoFill());
            chartShapePros.Append(outline);

            // Append c:spPr element to the end of c:chartSpace element
            chartSpace.Append(chartShapePros);

            chartPart.ChartSpace = chartSpace;

            // Generate content of the MainDocumentPart
            GeneratePartContent(mainPart);

            return document;

        }
        public static void GeneratePartContent(MainDocumentPart mainPart)
        {
            w.Paragraph paragraph = new w.Paragraph() { RsidParagraphAddition = "00C75AEB", RsidRunAdditionDefault = "000F3EFF" };

            // Create a new run that has an inline drawing object
            w.Run run = new w.Run();
            w.Drawing drawing = new w.Drawing();

            dw.Inline inline = new dw.Inline();
            inline.Append(new dw.Extent() { Cx = 5274310L, Cy = 3076575L });
            dw.DocProperties docPros = new dw.DocProperties() { Id = (UInt32Value)1U, Name = "Chart 1" };
            inline.Append(docPros);

            d.Graphic g = new d.Graphic();
            d.GraphicData graphicData = new d.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
            dc.ChartReference chartReference = new ChartReference() { Id = "rId110" };
            graphicData.Append(chartReference);
            g.Append(graphicData);
            inline.Append(g);
            drawing.Append(inline);
            run.Append(drawing);
            paragraph.Append(run);

            mainPart.Document.Body.Append(paragraph);
        }
    }
}
