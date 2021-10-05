using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace Generate_Word_Report.NewGen
{
    public class one_pic_format_position
    {
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId6");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId5");
            GenerateFontTablePart1Content(fontTablePart1);

            ImagePart imagePart1 = mainDocumentPart1.AddNewPart<ImagePart>("image/jpeg", "rId4");
            GenerateImagePart1Content(imagePart1);

        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "0";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "1";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "1";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "1";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            document1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            document1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            document1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            document1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            document1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            document1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            document1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            document1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            document1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            document1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            document1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "003F2413", RsidRunAdditionDefault = "00671749", ParagraphId = "70428F02", TextId = "10716AF0" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251658240U, BehindDoc = true, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "51BC60D6", AnchorId = "4502AFD8" };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Margin };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "2979300";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "1086641";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 2553335L, Cy = 3406775L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 3175L };

            Wp.WrapTight wrapTight1 = new Wp.WrapTight() { WrapText = Wp.WrapTextValues.BothSides };

            Wp.WrapPolygon wrapPolygon1 = new Wp.WrapPolygon() { Edited = false };
            Wp.StartPoint startPoint1 = new Wp.StartPoint() { X = 0L, Y = 0L };
            Wp.LineTo lineTo1 = new Wp.LineTo() { X = 0L, Y = 21499L };
            Wp.LineTo lineTo2 = new Wp.LineTo() { X = 21433L, Y = 21499L };
            Wp.LineTo lineTo3 = new Wp.LineTo() { X = 21433L, Y = 0L };
            Wp.LineTo lineTo4 = new Wp.LineTo() { X = 0L, Y = 0L };

            wrapPolygon1.Append(startPoint1);
            wrapPolygon1.Append(lineTo1);
            wrapPolygon1.Append(lineTo2);
            wrapPolygon1.Append(lineTo3);
            wrapPolygon1.Append(lineTo4);

            wrapTight1.Append(wrapPolygon1);
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 1" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId4", CompressionState = A.BlipCompressionValues.Print };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle();

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 2553335L, Cy = 3406775L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline1.Append(noFill2);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(outline1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapTight1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            run1.Append(runProperties1);
            run1.Append(drawing1);

            paragraph1.Append(run1);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "003F2413" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            webSettings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            webSettings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            webSettings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            webSettings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            webSettings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            AllowPNG allowPNG1 = new AllowPNG();

            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(allowPNG1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            settings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            settings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            settings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            settings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "110" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            Compatibility compatibility1 = new Compatibility();
            ApplyBreakingRules applyBreakingRules1 = new ApplyBreakingRules();
            UseFarEastLayout useFarEastLayout1 = new UseFarEastLayout();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting6 = new CompatibilitySetting() { Name = new EnumValue<CompatSettingNameValues>() { InnerText = "useWord2013TrackBottomHyphenation" }, Uri = "http://schemas.microsoft.com/office/word", Val = "0" };

            compatibility1.Append(applyBreakingRules1);
            compatibility1.Append(useFarEastLayout1);
            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);
            compatibility1.Append(compatibilitySetting6);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00745E8C" };
            Rsid rsid1 = new Rsid() { Val = "003E1962" };
            Rsid rsid2 = new Rsid() { Val = "003F2413" };
            Rsid rsid3 = new Rsid() { Val = "00671749" };
            Rsid rsid4 = new Rsid() { Val = "00745E8C" };
            Rsid rsid5 = new Rsid() { Val = "00AD54DD" };
            Rsid rsid6 = new Rsid() { Val = "00AF7273" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US", EastAsia = "ja-JP", Bidi = "th-TH" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{D3BA8806-B836-4991-B6CF-F04671E03FA6}" };

            settings1.Append(zoom1);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            styles1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            styles1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            styles1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "ja-JP", Bidi = "th-TH" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 376 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo372 = new LatentStyleExceptionInfo() { Name = "Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo373 = new LatentStyleExceptionInfo() { Name = "Smart Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo374 = new LatentStyleExceptionInfo() { Name = "Hashtag", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo375 = new LatentStyleExceptionInfo() { Name = "Unresolved Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo376 = new LatentStyleExceptionInfo() { Name = "Smart Link", SemiHidden = true, UnhideWhenUsed = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);
            latentStyles1.Append(latentStyleExceptionInfo372);
            latentStyles1.Append(latentStyleExceptionInfo373);
            latentStyles1.Append(latentStyleExceptionInfo374);
            latentStyles1.Append(latentStyleExceptionInfo375);
            latentStyles1.Append(latentStyleExceptionInfo376);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            style1.Append(styleName1);
            style1.Append(primaryStyle1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);
            style2.Append(unhideWhenUsed1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);
            majorFont1.Append(supplementalFont31);
            majorFont1.Append(supplementalFont32);
            majorFont1.Append(supplementalFont33);
            majorFont1.Append(supplementalFont34);
            majorFont1.Append(supplementalFont35);
            majorFont1.Append(supplementalFont36);
            majorFont1.Append(supplementalFont37);
            majorFont1.Append(supplementalFont38);
            majorFont1.Append(supplementalFont39);
            majorFont1.Append(supplementalFont40);
            majorFont1.Append(supplementalFont41);
            majorFont1.Append(supplementalFont42);
            majorFont1.Append(supplementalFont43);
            majorFont1.Append(supplementalFont44);
            majorFont1.Append(supplementalFont45);
            majorFont1.Append(supplementalFont46);
            majorFont1.Append(supplementalFont47);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游明朝" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont61 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont62 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont63 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont64 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont65 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont66 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont67 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont68 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont69 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont70 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont71 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont72 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont73 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont74 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont75 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont76 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont77 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont78 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont79 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont80 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont81 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont82 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont83 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont84 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont85 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont86 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont87 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont88 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont89 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont90 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont91 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont92 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont93 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont94 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);
            minorFont1.Append(supplementalFont61);
            minorFont1.Append(supplementalFont62);
            minorFont1.Append(supplementalFont63);
            minorFont1.Append(supplementalFont64);
            minorFont1.Append(supplementalFont65);
            minorFont1.Append(supplementalFont66);
            minorFont1.Append(supplementalFont67);
            minorFont1.Append(supplementalFont68);
            minorFont1.Append(supplementalFont69);
            minorFont1.Append(supplementalFont70);
            minorFont1.Append(supplementalFont71);
            minorFont1.Append(supplementalFont72);
            minorFont1.Append(supplementalFont73);
            minorFont1.Append(supplementalFont74);
            minorFont1.Append(supplementalFont75);
            minorFont1.Append(supplementalFont76);
            minorFont1.Append(supplementalFont77);
            minorFont1.Append(supplementalFont78);
            minorFont1.Append(supplementalFont79);
            minorFont1.Append(supplementalFont80);
            minorFont1.Append(supplementalFont81);
            minorFont1.Append(supplementalFont82);
            minorFont1.Append(supplementalFont83);
            minorFont1.Append(supplementalFont84);
            minorFont1.Append(supplementalFont85);
            minorFont1.Append(supplementalFont86);
            minorFont1.Append(supplementalFont87);
            minorFont1.Append(supplementalFont88);
            minorFont1.Append(supplementalFont89);
            minorFont1.Append(supplementalFont90);
            minorFont1.Append(supplementalFont91);
            minorFont1.Append(supplementalFont92);
            minorFont1.Append(supplementalFont93);
            minorFont1.Append(supplementalFont94);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline2 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill2);
            outline2.Append(presetDash1);
            outline2.Append(miter1);

            A.Outline outline3 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill3);
            outline3.Append(presetDash2);
            outline3.Append(miter2);

            A.Outline outline4 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline4.Append(solidFill4);
            outline4.Append(presetDash3);
            outline4.Append(miter3);

            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            fonts1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            fonts1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            fonts1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            fonts1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Yu Mincho" };
            AltName altName1 = new AltName() { Val = "游明朝" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020400000000000000" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "800002E7", UnicodeSignature1 = "2AC7FCFF", UnicodeSignature2 = "00000012", UnicodeSignature3 = "00000000", CodePageSignature0 = "0002009F", CodePageSignature1 = "00000000" };

            font2.Append(altName1);
            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Cordia New" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020B0304020202020204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "81000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00010001", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Yu Gothic Light" };
            AltName altName2 = new AltName() { Val = "游ゴシック Light" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020B0300000000000000" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "2AC7FDFF", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "0002009F", CodePageSignature1 = "00000000" };

            font5.Append(altName2);
            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            Font font7 = new Font() { Name = "Angsana New" };
            Panose1Number panose1Number7 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily7 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch7 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature7 = new FontSignature() { UnicodeSignature0 = "81000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00010001", CodePageSignature1 = "00000000" };

            font7.Append(panose1Number7);
            font7.Append(fontCharSet7);
            font7.Append(fontFamily7);
            font7.Append(pitch7);
            font7.Append(fontSignature7);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        

        #region Binary Data
        private string imagePart1Data = "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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}
