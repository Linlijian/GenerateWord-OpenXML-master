using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Generate_Word_Report
{
    class PicturePart2
    {
        private System.Collections.Generic.IDictionary<System.String, OpenXmlPart> UriPartDictionary = new System.Collections.Generic.Dictionary<System.String, OpenXmlPart>();
        private System.Collections.Generic.IDictionary<System.String, DataPart> UriNewDataPartDictionary = new System.Collections.Generic.Dictionary<System.String, DataPart>();
        private WordprocessingDocument document;

        public void ChangePackage(string filePath)
        {
            using (document = WordprocessingDocument.Open(filePath, true))
            {
                ChangeParts();
            }
        }

        private void ChangeParts()
        {
            //Stores the referrences to all the parts in a dictionary.
            BuildUriPartDictionary();
            //Changes the contents of the specified parts.
            ChangeExtendedFilePropertiesPart1(((ExtendedFilePropertiesPart)UriPartDictionary["/docProps/app.xml"]));
            ChangeCoreFilePropertiesPart1(((CoreFilePropertiesPart)UriPartDictionary["/docProps/core.xml"]));
            ChangeMainDocumentPart1(document.MainDocumentPart);
            ChangeDocumentSettingsPart1(((DocumentSettingsPart)UriPartDictionary["/word/settings.xml"]));
        }

        /// <summary>
        /// Stores the references to all the parts in the package.
        /// They could be retrieved by their URIs later.
        /// </summary>
        private void BuildUriPartDictionary()
        {
            System.Collections.Generic.Queue<OpenXmlPartContainer> queue = new System.Collections.Generic.Queue<OpenXmlPartContainer>();
            queue.Enqueue(document);
            while (queue.Count > 0)
            {
                foreach (var part in queue.Dequeue().Parts)
                {
                    if (!UriPartDictionary.Keys.Contains(part.OpenXmlPart.Uri.ToString()))
                    {
                        UriPartDictionary.Add(part.OpenXmlPart.Uri.ToString(), part.OpenXmlPart);
                        queue.Enqueue(part.OpenXmlPart);
                    }
                }
            }
        }

        private void ChangeExtendedFilePropertiesPart1(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = extendedFilePropertiesPart1.Properties;

            Ap.TotalTime totalTime1 = properties1.GetFirstChild<Ap.TotalTime>();
            Ap.Characters characters1 = properties1.GetFirstChild<Ap.Characters>();
            Ap.CharactersWithSpaces charactersWithSpaces1 = properties1.GetFirstChild<Ap.CharactersWithSpaces>();
            totalTime1.Text = "1";

            characters1.Text = "1";

            charactersWithSpaces1.Text = "1";

        }

        private void ChangeCoreFilePropertiesPart1(CoreFilePropertiesPart coreFilePropertiesPart1)
        {
            var package = coreFilePropertiesPart1.OpenXmlPackage;
            package.PackageProperties.Revision = "2";
            package.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-10-04T11:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            package.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2021-10-04T11:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        private void ChangeMainDocumentPart1(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = mainDocumentPart1.Document;

            Body body1 = document1.GetFirstChild<Body>();

            Paragraph paragraph1 = body1.GetFirstChild<Paragraph>();
            Paragraph paragraph2 = body1.Elements<Paragraph>().ElementAt(1);
            Paragraph paragraph3 = body1.Elements<Paragraph>().ElementAt(2);
            SectionProperties sectionProperties1 = body1.GetFirstChild<SectionProperties>();

            paragraph1.Remove();
            paragraph2.Remove();
            paragraph3.RsidParagraphProperties = "001B2F25";
            paragraph3.RsidRunAdditionDefault = "001B2F25";
            paragraph3.TextId = "7C3ADF94";
            paragraph3.RsidParagraphMarkRevision = "001B2F25";

            ParagraphProperties paragraphProperties1 = paragraph3.GetFirstChild<ParagraphProperties>();
            Run run1 = paragraph3.GetFirstChild<Run>();

            paragraphProperties1.Remove();

            RunProperties runProperties1 = run1.GetFirstChild<RunProperties>();
            Drawing drawing1 = run1.GetFirstChild<Drawing>();

            ComplexScript complexScript1 = runProperties1.GetFirstChild<ComplexScript>();

            complexScript1.Remove();

            Wp.Inline inline1 = drawing1.GetFirstChild<Wp.Inline>();
            inline1.AnchorId = "7050BAB0";
            inline1.EditId = "42FCEDC8";

            Wp.Extent extent1 = inline1.GetFirstChild<Wp.Extent>();
            Wp.EffectExtent effectExtent1 = inline1.GetFirstChild<Wp.EffectExtent>();
            Wp.DocProperties docProperties1 = inline1.GetFirstChild<Wp.DocProperties>();
            A.Graphic graphic1 = inline1.GetFirstChild<A.Graphic>();
            extent1.Cx = 5934710L;
            extent1.Cy = 7919085L;
            effectExtent1.RightEdge = 8890L;
            effectExtent1.BottomEdge = 5715L;
            docProperties1.Id = (UInt32Value)4U;
            docProperties1.Name = "Picture 4";

            A.GraphicData graphicData1 = graphic1.GetFirstChild<A.GraphicData>();

            Pic.Picture picture1 = graphicData1.GetFirstChild<Pic.Picture>();

            Pic.BlipFill blipFill1 = picture1.GetFirstChild<Pic.BlipFill>();
            Pic.ShapeProperties shapeProperties1 = picture1.GetFirstChild<Pic.ShapeProperties>();

            A.Blip blip1 = blipFill1.GetFirstChild<A.Blip>();
            blip1.CompressionState = null;

            A.Transform2D transform2D1 = shapeProperties1.GetFirstChild<A.Transform2D>();

            A.Extents extents1 = transform2D1.GetFirstChild<A.Extents>();
            extents1.Cx = 5934710L;
            extents1.Cy = 7919085L;
            sectionProperties1.RsidRPr = "001B2F25";
        }

        private void ChangeDocumentSettingsPart1(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = documentSettingsPart1.Settings;

            Rsids rsids1 = settings1.GetFirstChild<Rsids>();

            Rsid rsid1 = rsids1.GetFirstChild<Rsid>();

            Rsid rsid2 = new Rsid() { Val = "001B2F25" };
            rsids1.InsertBefore(rsid2, rsid1);
        }



    }
}
