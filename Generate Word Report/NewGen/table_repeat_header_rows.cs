using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace Generate_Word_Report.NewGen
{
    public class table_repeat_header_rows
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

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId1");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId6");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId5");
            GenerateFontTablePart1Content(fontTablePart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId4");
            GenerateWebSettingsPart1Content(webSettingsPart1);

        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "2";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "2";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "257";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "1466";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "12";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "3";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "1720";
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00561265", RsidParagraphProperties = "00B57ECF", RsidRunAdditionDefault = "00561265", ParagraphId = "38974014", TextId = "1E01BF42" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { Start = "1440" };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize1 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);

            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            paragraph1.Append(paragraphProperties1);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1451" };
            GridColumn gridColumn2 = new GridColumn() { Width = "5093" };
            GridColumn gridColumn3 = new GridColumn() { Width = "1878" };
            GridColumn gridColumn4 = new GridColumn() { Width = "1658" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "00561265", RsidTableRowProperties = "004D5E7B", ParagraphId = "24312B11", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableHeader tableHeader1 = new TableHeader();

            tableRowProperties1.Append(tableHeader1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin1 = new TableCellMargin();
            TopMargin topMargin1 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin1 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin1 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin1.Append(topMargin1);
            tableCellMargin1.Append(leftMargin1);
            tableCellMargin1.Append(bottomMargin1);
            tableCellMargin1.Append(rightMargin1);
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark1 = new HideMark();

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(verticalMerge1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellMargin1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);
            tableCellProperties1.Append(hideMark1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "1D821CC7", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize2 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties2.Append(runFonts2);
            paragraphMarkRunProperties2.Append(bold1);
            paragraphMarkRunProperties2.Append(boldComplexScript1);
            paragraphMarkRunProperties2.Append(fontSize2);

            paragraphProperties2.Append(spacingBetweenLines1);
            paragraphProperties2.Append(justification1);
            paragraphProperties2.Append(paragraphMarkRunProperties2);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize3 = new FontSize() { Val = "28" };

            runProperties1.Append(runFonts3);
            runProperties1.Append(bold2);
            runProperties1.Append(boldComplexScript2);
            runProperties1.Append(fontSize3);
            Text text1 = new Text();
            text1.Text = "องค์ประกอบหลัก";

            run1.Append(runProperties1);
            run1.Append(text1);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(proofError1);
            paragraph2.Append(run1);
            paragraph2.Append(proofError2);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4800", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge() { Val = MergedCellValues.Restart };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);
            tableCellBorders2.Append(leftBorder3);
            tableCellBorders2.Append(bottomBorder3);
            tableCellBorders2.Append(rightBorder3);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin2 = new TableCellMargin();
            TopMargin topMargin2 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin2 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin2 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin2.Append(topMargin2);
            tableCellMargin2.Append(leftMargin2);
            tableCellMargin2.Append(bottomMargin2);
            tableCellMargin2.Append(rightMargin2);
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark2 = new HideMark();

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(verticalMerge2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellMargin2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);
            tableCellProperties2.Append(hideMark2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "2396CC25", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0" };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize4 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties3.Append(runFonts4);
            paragraphMarkRunProperties3.Append(bold3);
            paragraphMarkRunProperties3.Append(boldComplexScript3);
            paragraphMarkRunProperties3.Append(fontSize4);

            paragraphProperties3.Append(spacingBetweenLines2);
            paragraphProperties3.Append(justification2);
            paragraphProperties3.Append(paragraphMarkRunProperties3);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            FontSize fontSize5 = new FontSize() { Val = "28" };

            runProperties2.Append(runFonts5);
            runProperties2.Append(bold4);
            runProperties2.Append(boldComplexScript4);
            runProperties2.Append(fontSize5);
            Text text2 = new Text();
            text2.Text = "รายละเอียดการวิเคราะห์";

            run2.Append(runProperties2);
            run2.Append(text2);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(proofError3);
            paragraph3.Append(run2);
            paragraph3.Append(proofError4);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph3);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 2 };

            TableCellBorders tableCellBorders3 = new TableCellBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders3.Append(topBorder4);
            tableCellBorders3.Append(leftBorder4);
            tableCellBorders3.Append(bottomBorder4);
            tableCellBorders3.Append(rightBorder4);
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin3 = new TableCellMargin();
            TopMargin topMargin3 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin3 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin3 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin3 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin3.Append(topMargin3);
            tableCellMargin3.Append(leftMargin3);
            tableCellMargin3.Append(bottomMargin3);
            tableCellMargin3.Append(rightMargin3);
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark3 = new HideMark();

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(gridSpan1);
            tableCellProperties3.Append(tableCellBorders3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellMargin3);
            tableCellProperties3.Append(tableCellVerticalAlignment3);
            tableCellProperties3.Append(hideMark3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "6B2276B7", TextId = "77777777" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0" };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold5 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            FontSize fontSize6 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties4.Append(runFonts6);
            paragraphMarkRunProperties4.Append(bold5);
            paragraphMarkRunProperties4.Append(boldComplexScript5);
            paragraphMarkRunProperties4.Append(fontSize6);

            paragraphProperties4.Append(spacingBetweenLines3);
            paragraphProperties4.Append(justification3);
            paragraphProperties4.Append(paragraphMarkRunProperties4);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold6 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            FontSize fontSize7 = new FontSize() { Val = "28" };

            runProperties3.Append(runFonts7);
            runProperties3.Append(bold6);
            runProperties3.Append(boldComplexScript6);
            runProperties3.Append(fontSize7);
            Text text3 = new Text();
            text3.Text = "องค์ประกอบย่อย";

            run3.Append(runProperties3);
            run3.Append(text3);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(proofError5);
            paragraph4.Append(run3);
            paragraph4.Append(proofError6);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph4);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "00561265", RsidTableRowProperties = "004D5E7B", ParagraphId = "71D73840", TextId = "77777777" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableHeader tableHeader2 = new TableHeader();

            tableRowProperties2.Append(tableHeader2);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge();

            TableCellBorders tableCellBorders4 = new TableCellBorders();
            TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders4.Append(topBorder5);
            tableCellBorders4.Append(leftBorder5);
            tableCellBorders4.Append(bottomBorder5);
            tableCellBorders4.Append(rightBorder5);
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin4 = new TableCellMargin();
            TopMargin topMargin4 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin4 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin4 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin4.Append(topMargin4);
            tableCellMargin4.Append(leftMargin4);
            tableCellMargin4.Append(bottomMargin4);
            tableCellMargin4.Append(rightMargin4);
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark4 = new HideMark();

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(verticalMerge3);
            tableCellProperties4.Append(tableCellBorders4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellMargin4);
            tableCellProperties4.Append(tableCellVerticalAlignment4);
            tableCellProperties4.Append(hideMark4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00561265", ParagraphId = "0D4DE3E6", TextId = "77777777" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            FontSize fontSize8 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties5.Append(fontSize8);

            paragraphProperties5.Append(paragraphMarkRunProperties5);

            paragraph5.Append(paragraphProperties5);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph5);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "4800", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge();

            TableCellBorders tableCellBorders5 = new TableCellBorders();
            TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders5.Append(topBorder6);
            tableCellBorders5.Append(leftBorder6);
            tableCellBorders5.Append(bottomBorder6);
            tableCellBorders5.Append(rightBorder6);
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin5 = new TableCellMargin();
            TopMargin topMargin5 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin5 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin5 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin5 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin5.Append(topMargin5);
            tableCellMargin5.Append(leftMargin5);
            tableCellMargin5.Append(bottomMargin5);
            tableCellMargin5.Append(rightMargin5);
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark5 = new HideMark();

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(verticalMerge4);
            tableCellProperties5.Append(tableCellBorders5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellMargin5);
            tableCellProperties5.Append(tableCellVerticalAlignment5);
            tableCellProperties5.Append(hideMark5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00561265", ParagraphId = "002CB54B", TextId = "77777777" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            FontSize fontSize9 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties6.Append(fontSize9);

            paragraphProperties6.Append(paragraphMarkRunProperties6);

            paragraph6.Append(paragraphProperties6);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph6);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders6 = new TableCellBorders();
            TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders6.Append(topBorder7);
            tableCellBorders6.Append(leftBorder7);
            tableCellBorders6.Append(bottomBorder7);
            tableCellBorders6.Append(rightBorder7);
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin6 = new TableCellMargin();
            TopMargin topMargin6 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin6 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin6 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin6 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin6.Append(topMargin6);
            tableCellMargin6.Append(leftMargin6);
            tableCellMargin6.Append(bottomMargin6);
            tableCellMargin6.Append(rightMargin6);
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark6 = new HideMark();

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(tableCellBorders6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellMargin6);
            tableCellProperties6.Append(tableCellVerticalAlignment6);
            tableCellProperties6.Append(hideMark6);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "4E339D57", TextId = "77777777" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "0" };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold7 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            FontSize fontSize10 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties7.Append(runFonts8);
            paragraphMarkRunProperties7.Append(bold7);
            paragraphMarkRunProperties7.Append(boldComplexScript7);
            paragraphMarkRunProperties7.Append(fontSize10);

            paragraphProperties7.Append(spacingBetweenLines4);
            paragraphProperties7.Append(justification4);
            paragraphProperties7.Append(paragraphMarkRunProperties7);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold8 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            FontSize fontSize11 = new FontSize() { Val = "28" };

            runProperties4.Append(runFonts9);
            runProperties4.Append(bold8);
            runProperties4.Append(boldComplexScript8);
            runProperties4.Append(fontSize11);
            Text text4 = new Text();
            text4.Text = "บริษัทประกันวินาศภัย";

            run4.Append(runProperties4);
            run4.Append(text4);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(proofError7);
            paragraph7.Append(run4);
            paragraph7.Append(proofError8);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph7);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders7 = new TableCellBorders();
            TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders7.Append(topBorder8);
            tableCellBorders7.Append(leftBorder8);
            tableCellBorders7.Append(bottomBorder8);
            tableCellBorders7.Append(rightBorder8);
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin7 = new TableCellMargin();
            TopMargin topMargin7 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin7 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin7 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin7 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin7.Append(topMargin7);
            tableCellMargin7.Append(leftMargin7);
            tableCellMargin7.Append(bottomMargin7);
            tableCellMargin7.Append(rightMargin7);
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark7 = new HideMark();

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(tableCellBorders7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellMargin7);
            tableCellProperties7.Append(tableCellVerticalAlignment7);
            tableCellProperties7.Append(hideMark7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "2D27FB0F", TextId = "77777777" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { After = "0" };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold9 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            FontSize fontSize12 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties8.Append(runFonts10);
            paragraphMarkRunProperties8.Append(bold9);
            paragraphMarkRunProperties8.Append(boldComplexScript9);
            paragraphMarkRunProperties8.Append(fontSize12);

            paragraphProperties8.Append(spacingBetweenLines5);
            paragraphProperties8.Append(justification5);
            paragraphProperties8.Append(paragraphMarkRunProperties8);
            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run5 = new Run();

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            Bold bold10 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            FontSize fontSize13 = new FontSize() { Val = "28" };

            runProperties5.Append(runFonts11);
            runProperties5.Append(bold10);
            runProperties5.Append(boldComplexScript10);
            runProperties5.Append(fontSize13);
            Text text5 = new Text();
            text5.Text = "บริษัทประกันชีวิต";

            run5.Append(runProperties5);
            run5.Append(text5);
            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(proofError9);
            paragraph8.Append(run5);
            paragraph8.Append(proofError10);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph8);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);
            tableRow2.Append(tableCell7);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "00561265", ParagraphId = "0FD025F7", TextId = "77777777" };

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders8 = new TableCellBorders();
            TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders8.Append(topBorder9);
            tableCellBorders8.Append(leftBorder9);
            tableCellBorders8.Append(bottomBorder9);
            tableCellBorders8.Append(rightBorder9);
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin8 = new TableCellMargin();
            TopMargin topMargin8 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin8 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin8 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin8 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin8.Append(topMargin8);
            tableCellMargin8.Append(leftMargin8);
            tableCellMargin8.Append(bottomMargin8);
            tableCellMargin8.Append(rightMargin8);
            HideMark hideMark8 = new HideMark();

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellBorders8);
            tableCellProperties8.Append(shading8);
            tableCellProperties8.Append(tableCellMargin8);
            tableCellProperties8.Append(hideMark8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "6895EC3A", TextId = "77777777" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize14 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties9.Append(runFonts12);
            paragraphMarkRunProperties9.Append(fontSize14);

            paragraphProperties9.Append(spacingBetweenLines6);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize15 = new FontSize() { Val = "28" };

            runProperties6.Append(runFonts13);
            runProperties6.Append(fontSize15);
            Text text6 = new Text();
            text6.Text = "1. Profitability";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run6);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph9);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "4800", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders9.Append(topBorder10);
            tableCellBorders9.Append(leftBorder10);
            tableCellBorders9.Append(bottomBorder10);
            tableCellBorders9.Append(rightBorder10);
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin9 = new TableCellMargin();
            TopMargin topMargin9 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin9 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin9 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin9 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin9.Append(topMargin9);
            tableCellMargin9.Append(leftMargin9);
            tableCellMargin9.Append(bottomMargin9);
            tableCellMargin9.Append(rightMargin9);
            HideMark hideMark9 = new HideMark();

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellBorders9);
            tableCellProperties9.Append(shading9);
            tableCellProperties9.Append(tableCellMargin9);
            tableCellProperties9.Append(hideMark9);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "5A946C5F", TextId = "77777777" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize16 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties10.Append(runFonts14);
            paragraphMarkRunProperties10.Append(fontSize16);

            paragraphProperties10.Append(spacingBetweenLines7);
            paragraphProperties10.Append(paragraphMarkRunProperties10);
            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize17 = new FontSize() { Val = "28" };

            runProperties7.Append(runFonts15);
            runProperties7.Append(fontSize17);
            Text text7 = new Text();
            text7.Text = "การวิเคราะห์พอร์ตการรับประกันภัยและผลการดำเนินงาน";

            run7.Append(runProperties7);
            run7.Append(text7);
            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize18 = new FontSize() { Val = "28" };

            runProperties8.Append(runFonts16);
            runProperties8.Append(fontSize18);
            Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text8.Text = " ";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(proofError11);
            paragraph10.Append(run7);
            paragraph10.Append(proofError12);
            paragraph10.Append(run8);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "0B0E69A8", TextId = "77777777" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize19 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties11.Append(runFonts17);
            paragraphMarkRunProperties11.Append(fontSize19);

            paragraphProperties11.Append(spacingBetweenLines8);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize20 = new FontSize() { Val = "28" };

            runProperties9.Append(runFonts18);
            runProperties9.Append(fontSize20);
            Text text9 = new Text();
            text9.Text = "(";

            run9.Append(runProperties9);
            run9.Append(text9);
            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize21 = new FontSize() { Val = "28" };

            runProperties10.Append(runFonts19);
            runProperties10.Append(fontSize21);
            Text text10 = new Text();
            text10.Text = "เบี้ยประกันภัยค่าสินไหมทดแทนและค่าใช้จ่ายของบริษัท";

            run10.Append(runProperties10);
            run10.Append(text10);
            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize22 = new FontSize() { Val = "28" };

            runProperties11.Append(runFonts20);
            runProperties11.Append(fontSize22);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = ") ";

            run11.Append(runProperties11);
            run11.Append(text11);
            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run12 = new Run();

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize23 = new FontSize() { Val = "28" };

            runProperties12.Append(runFonts21);
            runProperties12.Append(fontSize23);
            Text text12 = new Text();
            text12.Text = "รวมถึง";

            run12.Append(runProperties12);
            run12.Append(text12);
            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize24 = new FontSize() { Val = "28" };

            runProperties13.Append(runFonts22);
            runProperties13.Append(fontSize24);
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = " ";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run9);
            paragraph11.Append(proofError13);
            paragraph11.Append(run10);
            paragraph11.Append(proofError14);
            paragraph11.Append(run11);
            paragraph11.Append(proofError15);
            paragraph11.Append(run12);
            paragraph11.Append(proofError16);
            paragraph11.Append(run13);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "42F81026", TextId = "77777777" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize25 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties12.Append(runFonts23);
            paragraphMarkRunProperties12.Append(fontSize25);

            paragraphProperties12.Append(spacingBetweenLines9);
            paragraphProperties12.Append(paragraphMarkRunProperties12);
            ProofError proofError17 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize26 = new FontSize() { Val = "28" };

            runProperties14.Append(runFonts24);
            runProperties14.Append(fontSize26);
            Text text14 = new Text();
            text14.Text = "ปัจจัยความเสี่ยงที่เกี่ยวข้องกับการรับประกันภัย";

            run14.Append(runProperties14);
            run14.Append(text14);
            ProofError proofError18 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize27 = new FontSize() { Val = "28" };

            runProperties15.Append(runFonts25);
            runProperties15.Append(fontSize27);
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = " ";

            run15.Append(runProperties15);
            run15.Append(text15);
            ProofError proofError19 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run16 = new Run();

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize28 = new FontSize() { Val = "28" };

            runProperties16.Append(runFonts26);
            runProperties16.Append(fontSize28);
            Text text16 = new Text();
            text16.Text = "อาทิ";

            run16.Append(runProperties16);
            run16.Append(text16);
            ProofError proofError20 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run17 = new Run();

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize29 = new FontSize() { Val = "28" };

            runProperties17.Append(runFonts27);
            runProperties17.Append(fontSize29);
            Text text17 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text17.Text = " ";

            run17.Append(runProperties17);
            run17.Append(text17);
            ProofError proofError21 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run18 = new Run();

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize30 = new FontSize() { Val = "28" };

            runProperties18.Append(runFonts28);
            runProperties18.Append(fontSize30);
            Text text18 = new Text();
            text18.Text = "ช่องทาง";

            run18.Append(runProperties18);
            run18.Append(text18);
            ProofError proofError22 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(proofError17);
            paragraph12.Append(run14);
            paragraph12.Append(proofError18);
            paragraph12.Append(run15);
            paragraph12.Append(proofError19);
            paragraph12.Append(run16);
            paragraph12.Append(proofError20);
            paragraph12.Append(run17);
            paragraph12.Append(proofError21);
            paragraph12.Append(run18);
            paragraph12.Append(proofError22);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "41904B89", TextId = "77777777" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize31 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties13.Append(runFonts29);
            paragraphMarkRunProperties13.Append(fontSize31);

            paragraphProperties13.Append(spacingBetweenLines10);
            paragraphProperties13.Append(paragraphMarkRunProperties13);
            ProofError proofError23 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run19 = new Run();

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize32 = new FontSize() { Val = "28" };

            runProperties19.Append(runFonts30);
            runProperties19.Append(fontSize32);
            Text text19 = new Text();
            text19.Text = "การขายและการติดตามเบี้ยประกันภัยค้างรับ";

            run19.Append(runProperties19);
            run19.Append(text19);
            ProofError proofError24 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run20 = new Run();

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize33 = new FontSize() { Val = "28" };

            runProperties20.Append(runFonts31);
            runProperties20.Append(fontSize33);
            Text text20 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text20.Text = " ";

            run20.Append(runProperties20);
            run20.Append(text20);
            ProofError proofError25 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run21 = new Run();

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize34 = new FontSize() { Val = "28" };

            runProperties21.Append(runFonts32);
            runProperties21.Append(fontSize34);
            Text text21 = new Text();
            text21.Text = "เป็นต้น";

            run21.Append(runProperties21);
            run21.Append(text21);
            ProofError proofError26 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(proofError23);
            paragraph13.Append(run19);
            paragraph13.Append(proofError24);
            paragraph13.Append(run20);
            paragraph13.Append(proofError25);
            paragraph13.Append(run21);
            paragraph13.Append(proofError26);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph10);
            tableCell9.Append(paragraph11);
            tableCell9.Append(paragraph12);
            tableCell9.Append(paragraph13);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders10.Append(topBorder11);
            tableCellBorders10.Append(leftBorder11);
            tableCellBorders10.Append(bottomBorder11);
            tableCellBorders10.Append(rightBorder11);
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin10 = new TableCellMargin();
            TopMargin topMargin10 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin10 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin10 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin10 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin10.Append(topMargin10);
            tableCellMargin10.Append(leftMargin10);
            tableCellMargin10.Append(bottomMargin10);
            tableCellMargin10.Append(rightMargin10);
            HideMark hideMark10 = new HideMark();

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(tableCellBorders10);
            tableCellProperties10.Append(shading10);
            tableCellProperties10.Append(tableCellMargin10);
            tableCellProperties10.Append(hideMark10);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "57AE966B", TextId = "77777777" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize35 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties14.Append(runFonts33);
            paragraphMarkRunProperties14.Append(fontSize35);

            paragraphProperties14.Append(spacingBetweenLines11);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run22 = new Run();

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize36 = new FontSize() { Val = "28" };

            runProperties22.Append(runFonts34);
            runProperties22.Append(fontSize36);
            Text text22 = new Text();
            text22.Text = "- Loss ratio";

            run22.Append(runProperties22);
            run22.Append(text22);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run22);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "5FCD63EE", TextId = "77777777" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize37 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties15.Append(runFonts35);
            paragraphMarkRunProperties15.Append(fontSize37);

            paragraphProperties15.Append(spacingBetweenLines12);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run23 = new Run();

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize38 = new FontSize() { Val = "28" };

            runProperties23.Append(runFonts36);
            runProperties23.Append(fontSize38);
            Text text23 = new Text();
            text23.Text = "- Expense ratio";

            run23.Append(runProperties23);
            run23.Append(text23);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run23);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "657F923A", TextId = "77777777" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize39 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties16.Append(runFonts37);
            paragraphMarkRunProperties16.Append(fontSize39);

            paragraphProperties16.Append(spacingBetweenLines13);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run24 = new Run();

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize40 = new FontSize() { Val = "28" };

            runProperties24.Append(runFonts38);
            runProperties24.Append(fontSize40);
            Text text24 = new Text();
            text24.Text = "- Premium receivable before impairment";

            run24.Append(runProperties24);
            run24.Append(text24);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run24);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "6A19A312", TextId = "77777777" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize41 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties17.Append(runFonts39);
            paragraphMarkRunProperties17.Append(fontSize41);

            paragraphProperties17.Append(spacingBetweenLines14);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run25 = new Run();

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize42 = new FontSize() { Val = "28" };

            runProperties25.Append(runFonts40);
            runProperties25.Append(fontSize42);
            Text text25 = new Text();
            text25.Text = "- Return on equity (ROE)";

            run25.Append(runProperties25);
            run25.Append(text25);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run25);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph14);
            tableCell10.Append(paragraph15);
            tableCell10.Append(paragraph16);
            tableCell10.Append(paragraph17);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "1900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            TopBorder topBorder12 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder12 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders11.Append(topBorder12);
            tableCellBorders11.Append(leftBorder12);
            tableCellBorders11.Append(bottomBorder12);
            tableCellBorders11.Append(rightBorder12);
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin11 = new TableCellMargin();
            TopMargin topMargin11 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin11 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin11 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin11 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin11.Append(topMargin11);
            tableCellMargin11.Append(leftMargin11);
            tableCellMargin11.Append(bottomMargin11);
            tableCellMargin11.Append(rightMargin11);
            HideMark hideMark11 = new HideMark();

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(tableCellBorders11);
            tableCellProperties11.Append(shading11);
            tableCellProperties11.Append(tableCellMargin11);
            tableCellProperties11.Append(hideMark11);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "7491C7DF", TextId = "77777777" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize43 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties18.Append(runFonts41);
            paragraphMarkRunProperties18.Append(fontSize43);

            paragraphProperties18.Append(spacingBetweenLines15);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run26 = new Run();

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize44 = new FontSize() { Val = "28" };

            runProperties26.Append(runFonts42);
            runProperties26.Append(fontSize44);
            Text text26 = new Text();
            text26.Text = "- Expense ratio";

            run26.Append(runProperties26);
            run26.Append(text26);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run26);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "17ADAA27", TextId = "77777777" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize45 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties19.Append(runFonts43);
            paragraphMarkRunProperties19.Append(fontSize45);

            paragraphProperties19.Append(spacingBetweenLines16);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run27 = new Run();

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize46 = new FontSize() { Val = "28" };

            runProperties27.Append(runFonts44);
            runProperties27.Append(fontSize46);
            Text text27 = new Text();
            text27.Text = "- Change in net written premium";

            run27.Append(runProperties27);
            run27.Append(text27);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run27);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "7C65629F", TextId = "77777777" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize47 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties20.Append(runFonts45);
            paragraphMarkRunProperties20.Append(fontSize47);

            paragraphProperties20.Append(spacingBetweenLines17);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run28 = new Run();

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize48 = new FontSize() { Val = "28" };

            runProperties28.Append(runFonts46);
            runProperties28.Append(fontSize48);
            Text text28 = new Text();
            text28.Text = "- Change in single premium";

            run28.Append(runProperties28);
            run28.Append(text28);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run28);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "58856067", TextId = "77777777" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize49 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties21.Append(runFonts47);
            paragraphMarkRunProperties21.Append(fontSize49);

            paragraphProperties21.Append(spacingBetweenLines18);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run29 = new Run();

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize50 = new FontSize() { Val = "28" };

            runProperties29.Append(runFonts48);
            runProperties29.Append(fontSize50);
            Text text29 = new Text();
            text29.Text = "- Return on equity (ROE)";

            run29.Append(runProperties29);
            run29.Append(text29);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run29);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph18);
            tableCell11.Append(paragraph19);
            tableCell11.Append(paragraph20);
            tableCell11.Append(paragraph21);

            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);

            TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "00561265", ParagraphId = "3EDB95EE", TextId = "77777777" };

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders12 = new TableCellBorders();
            TopBorder topBorder13 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder13 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders12.Append(topBorder13);
            tableCellBorders12.Append(leftBorder13);
            tableCellBorders12.Append(bottomBorder13);
            tableCellBorders12.Append(rightBorder13);
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin12 = new TableCellMargin();
            TopMargin topMargin12 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin12 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin12 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin12 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin12.Append(topMargin12);
            tableCellMargin12.Append(leftMargin12);
            tableCellMargin12.Append(bottomMargin12);
            tableCellMargin12.Append(rightMargin12);
            HideMark hideMark12 = new HideMark();

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(tableCellBorders12);
            tableCellProperties12.Append(shading12);
            tableCellProperties12.Append(tableCellMargin12);
            tableCellProperties12.Append(hideMark12);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "3F4D2057", TextId = "77777777" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize51 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties22.Append(runFonts49);
            paragraphMarkRunProperties22.Append(fontSize51);

            paragraphProperties22.Append(spacingBetweenLines19);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run30 = new Run();

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize52 = new FontSize() { Val = "28" };

            runProperties30.Append(runFonts50);
            runProperties30.Append(fontSize52);
            Text text30 = new Text();
            text30.Text = "2. Capital adequacy";

            run30.Append(runProperties30);
            run30.Append(text30);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run30);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph22);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "4800", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders13 = new TableCellBorders();
            TopBorder topBorder14 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder14 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder14 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders13.Append(topBorder14);
            tableCellBorders13.Append(leftBorder14);
            tableCellBorders13.Append(bottomBorder14);
            tableCellBorders13.Append(rightBorder14);
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin13 = new TableCellMargin();
            TopMargin topMargin13 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin13 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin13 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin13 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin13.Append(topMargin13);
            tableCellMargin13.Append(leftMargin13);
            tableCellMargin13.Append(bottomMargin13);
            tableCellMargin13.Append(rightMargin13);
            HideMark hideMark13 = new HideMark();

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellBorders13);
            tableCellProperties13.Append(shading13);
            tableCellProperties13.Append(tableCellMargin13);
            tableCellProperties13.Append(hideMark13);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "0CDE7882", TextId = "77777777" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize53 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties23.Append(runFonts51);
            paragraphMarkRunProperties23.Append(fontSize53);

            paragraphProperties23.Append(spacingBetweenLines20);
            paragraphProperties23.Append(paragraphMarkRunProperties23);
            ProofError proofError27 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run31 = new Run();

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize54 = new FontSize() { Val = "28" };

            runProperties31.Append(runFonts52);
            runProperties31.Append(fontSize54);
            Text text31 = new Text();
            text31.Text = "การวิเคราะห์องค์ประกอบของเงินกองทุนที่สามารถนำมาใช้ได้ทั้งหมด";

            run31.Append(runProperties31);
            run31.Append(text31);
            ProofError proofError28 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run32 = new Run();

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize55 = new FontSize() { Val = "28" };

            runProperties32.Append(runFonts53);
            runProperties32.Append(fontSize55);
            Text text32 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text32.Text = " ";

            run32.Append(runProperties32);
            run32.Append(text32);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(proofError27);
            paragraph23.Append(run31);
            paragraph23.Append(proofError28);
            paragraph23.Append(run32);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "4DC1D0F8", TextId = "77777777" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize56 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties24.Append(runFonts54);
            paragraphMarkRunProperties24.Append(fontSize56);

            paragraphProperties24.Append(spacingBetweenLines21);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run33 = new Run();

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize57 = new FontSize() { Val = "28" };

            runProperties33.Append(runFonts55);
            runProperties33.Append(fontSize57);
            Text text33 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text33.Text = "(TCA) ";

            run33.Append(runProperties33);
            run33.Append(text33);
            ProofError proofError29 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run34 = new Run();

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize58 = new FontSize() { Val = "28" };

            runProperties34.Append(runFonts56);
            runProperties34.Append(fontSize58);
            Text text34 = new Text();
            text34.Text = "และเงินกองทุนที่ต้องดำรงทั้งหมด";

            run34.Append(runProperties34);
            run34.Append(text34);
            ProofError proofError30 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run35 = new Run();

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize59 = new FontSize() { Val = "28" };

            runProperties35.Append(runFonts57);
            runProperties35.Append(fontSize59);
            Text text35 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text35.Text = " (TCR) ";

            run35.Append(runProperties35);
            run35.Append(text35);
            ProofError proofError31 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run36 = new Run();

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize60 = new FontSize() { Val = "28" };

            runProperties36.Append(runFonts58);
            runProperties36.Append(fontSize60);
            Text text36 = new Text();
            text36.Text = "ซึ่งต้องสามารถ";

            run36.Append(runProperties36);
            run36.Append(text36);
            ProofError proofError32 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run33);
            paragraph24.Append(proofError29);
            paragraph24.Append(run34);
            paragraph24.Append(proofError30);
            paragraph24.Append(run35);
            paragraph24.Append(proofError31);
            paragraph24.Append(run36);
            paragraph24.Append(proofError32);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "1F7A8759", TextId = "77777777" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts59 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize61 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties25.Append(runFonts59);
            paragraphMarkRunProperties25.Append(fontSize61);

            paragraphProperties25.Append(spacingBetweenLines22);
            paragraphProperties25.Append(paragraphMarkRunProperties25);
            ProofError proofError33 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run37 = new Run();

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize62 = new FontSize() { Val = "28" };

            runProperties37.Append(runFonts60);
            runProperties37.Append(fontSize62);
            Text text37 = new Text();
            text37.Text = "แสดงให้เห็นถึงปัจจัยเสี่ยงที่กระทบต่อระดับเงินกองทุนของบริษัท";

            run37.Append(runProperties37);
            run37.Append(text37);
            ProofError proofError34 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph25.Append(paragraphProperties25);
            paragraph25.Append(proofError33);
            paragraph25.Append(run37);
            paragraph25.Append(proofError34);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph23);
            tableCell13.Append(paragraph24);
            tableCell13.Append(paragraph25);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders14 = new TableCellBorders();
            TopBorder topBorder15 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder15 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder15 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders14.Append(topBorder15);
            tableCellBorders14.Append(leftBorder15);
            tableCellBorders14.Append(bottomBorder15);
            tableCellBorders14.Append(rightBorder15);
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin14 = new TableCellMargin();
            TopMargin topMargin14 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin14 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin14 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin14 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin14.Append(topMargin14);
            tableCellMargin14.Append(leftMargin14);
            tableCellMargin14.Append(bottomMargin14);
            tableCellMargin14.Append(rightMargin14);
            HideMark hideMark14 = new HideMark();

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(tableCellBorders14);
            tableCellProperties14.Append(shading14);
            tableCellProperties14.Append(tableCellMargin14);
            tableCellProperties14.Append(hideMark14);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "3AA71F39", TextId = "77777777" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize63 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties26.Append(runFonts61);
            paragraphMarkRunProperties26.Append(fontSize63);

            paragraphProperties26.Append(spacingBetweenLines23);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            Run run38 = new Run();

            RunProperties runProperties38 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize64 = new FontSize() { Val = "28" };

            runProperties38.Append(runFonts62);
            runProperties38.Append(fontSize64);
            Text text38 = new Text();
            text38.Text = "- Capital adequacy ratio (CAR)";

            run38.Append(runProperties38);
            run38.Append(text38);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(run38);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "3865A8E5", TextId = "77777777" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines24 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize65 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties27.Append(runFonts63);
            paragraphMarkRunProperties27.Append(fontSize65);

            paragraphProperties27.Append(spacingBetweenLines24);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            Run run39 = new Run();

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize66 = new FontSize() { Val = "28" };

            runProperties39.Append(runFonts64);
            runProperties39.Append(fontSize66);
            Text text39 = new Text();
            text39.Text = "- Change in TCA";

            run39.Append(runProperties39);
            run39.Append(text39);

            paragraph27.Append(paragraphProperties27);
            paragraph27.Append(run39);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "2F0D3534", TextId = "77777777" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize67 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties28.Append(runFonts65);
            paragraphMarkRunProperties28.Append(fontSize67);

            paragraphProperties28.Append(spacingBetweenLines25);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            Run run40 = new Run();

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize68 = new FontSize() { Val = "28" };

            runProperties40.Append(runFonts66);
            runProperties40.Append(fontSize68);
            Text text40 = new Text();
            text40.Text = "- Net written premium per TCA";

            run40.Append(runProperties40);
            run40.Append(text40);

            paragraph28.Append(paragraphProperties28);
            paragraph28.Append(run40);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "28BEE7D4", TextId = "77777777" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize69 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties29.Append(runFonts67);
            paragraphMarkRunProperties29.Append(fontSize69);

            paragraphProperties29.Append(spacingBetweenLines26);
            paragraphProperties29.Append(paragraphMarkRunProperties29);

            Run run41 = new Run();

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize70 = new FontSize() { Val = "28" };

            runProperties41.Append(runFonts68);
            runProperties41.Append(fontSize70);
            Text text41 = new Text();
            text41.Text = "- Commission income per TCA";

            run41.Append(runProperties41);
            run41.Append(text41);

            paragraph29.Append(paragraphProperties29);
            paragraph29.Append(run41);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph26);
            tableCell14.Append(paragraph27);
            tableCell14.Append(paragraph28);
            tableCell14.Append(paragraph29);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders15 = new TableCellBorders();
            TopBorder topBorder16 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder16 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder16 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders15.Append(topBorder16);
            tableCellBorders15.Append(leftBorder16);
            tableCellBorders15.Append(bottomBorder16);
            tableCellBorders15.Append(rightBorder16);
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin15 = new TableCellMargin();
            TopMargin topMargin15 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin15 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin15 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin15 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin15.Append(topMargin15);
            tableCellMargin15.Append(leftMargin15);
            tableCellMargin15.Append(bottomMargin15);
            tableCellMargin15.Append(rightMargin15);
            HideMark hideMark15 = new HideMark();

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(tableCellBorders15);
            tableCellProperties15.Append(shading15);
            tableCellProperties15.Append(tableCellMargin15);
            tableCellProperties15.Append(hideMark15);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "3ED084FC", TextId = "77777777" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines27 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts69 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize71 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties30.Append(runFonts69);
            paragraphMarkRunProperties30.Append(fontSize71);

            paragraphProperties30.Append(spacingBetweenLines27);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            Run run42 = new Run();

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize72 = new FontSize() { Val = "28" };

            runProperties42.Append(runFonts70);
            runProperties42.Append(fontSize72);
            Text text42 = new Text();
            text42.Text = "- Capital adequacy ratio (CAR)";

            run42.Append(runProperties42);
            run42.Append(text42);

            paragraph30.Append(paragraphProperties30);
            paragraph30.Append(run42);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "608A90A3", TextId = "77777777" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines28 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize73 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties31.Append(runFonts71);
            paragraphMarkRunProperties31.Append(fontSize73);

            paragraphProperties31.Append(spacingBetweenLines28);
            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run43 = new Run();

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize74 = new FontSize() { Val = "28" };

            runProperties43.Append(runFonts72);
            runProperties43.Append(fontSize74);
            Text text43 = new Text();
            text43.Text = "- Change in TCA";

            run43.Append(runProperties43);
            run43.Append(text43);

            paragraph31.Append(paragraphProperties31);
            paragraph31.Append(run43);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph30);
            tableCell15.Append(paragraph31);

            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);
            tableRow4.Append(tableCell14);
            tableRow4.Append(tableCell15);

            TableRow tableRow5 = new TableRow() { RsidTableRowAddition = "00561265", ParagraphId = "437DC8B5", TextId = "77777777" };

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders16 = new TableCellBorders();
            TopBorder topBorder17 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder17 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder17 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder17 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders16.Append(topBorder17);
            tableCellBorders16.Append(leftBorder17);
            tableCellBorders16.Append(bottomBorder17);
            tableCellBorders16.Append(rightBorder17);
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin16 = new TableCellMargin();
            TopMargin topMargin16 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin16 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin16 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin16 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin16.Append(topMargin16);
            tableCellMargin16.Append(leftMargin16);
            tableCellMargin16.Append(bottomMargin16);
            tableCellMargin16.Append(rightMargin16);
            HideMark hideMark16 = new HideMark();

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(tableCellBorders16);
            tableCellProperties16.Append(shading16);
            tableCellProperties16.Append(tableCellMargin16);
            tableCellProperties16.Append(hideMark16);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "0913A7B4", TextId = "77777777" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines29 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts73 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize75 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties32.Append(runFonts73);
            paragraphMarkRunProperties32.Append(fontSize75);

            paragraphProperties32.Append(spacingBetweenLines29);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            Run run44 = new Run();

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts74 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize76 = new FontSize() { Val = "28" };

            runProperties44.Append(runFonts74);
            runProperties44.Append(fontSize76);
            Text text44 = new Text();
            text44.Text = "3. Liquidity";

            run44.Append(runProperties44);
            run44.Append(text44);

            paragraph32.Append(paragraphProperties32);
            paragraph32.Append(run44);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph32);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "4800", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders17 = new TableCellBorders();
            TopBorder topBorder18 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder18 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder18 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder18 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders17.Append(topBorder18);
            tableCellBorders17.Append(leftBorder18);
            tableCellBorders17.Append(bottomBorder18);
            tableCellBorders17.Append(rightBorder18);
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin17 = new TableCellMargin();
            TopMargin topMargin17 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin17 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin17 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin17 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin17.Append(topMargin17);
            tableCellMargin17.Append(leftMargin17);
            tableCellMargin17.Append(bottomMargin17);
            tableCellMargin17.Append(rightMargin17);
            HideMark hideMark17 = new HideMark();

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellBorders17);
            tableCellProperties17.Append(shading17);
            tableCellProperties17.Append(tableCellMargin17);
            tableCellProperties17.Append(hideMark17);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "52154881", TextId = "77777777" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines30 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts75 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize77 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties33.Append(runFonts75);
            paragraphMarkRunProperties33.Append(fontSize77);

            paragraphProperties33.Append(spacingBetweenLines30);
            paragraphProperties33.Append(paragraphMarkRunProperties33);
            ProofError proofError35 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run45 = new Run();

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts76 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize78 = new FontSize() { Val = "28" };

            runProperties45.Append(runFonts76);
            runProperties45.Append(fontSize78);
            Text text45 = new Text();
            text45.Text = "การวิเคราะห์กระแสเงินสดรับ-จ่ายของบริษัท";

            run45.Append(runProperties45);
            run45.Append(text45);
            ProofError proofError36 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run46 = new Run();

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize79 = new FontSize() { Val = "28" };

            runProperties46.Append(runFonts77);
            runProperties46.Append(fontSize79);
            Text text46 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text46.Text = " ";

            run46.Append(runProperties46);
            run46.Append(text46);
            ProofError proofError37 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run47 = new Run();

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts78 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize80 = new FontSize() { Val = "28" };

            runProperties47.Append(runFonts78);
            runProperties47.Append(fontSize80);
            Text text47 = new Text();
            text47.Text = "และวัดความสามารถ";

            run47.Append(runProperties47);
            run47.Append(text47);
            ProofError proofError38 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph33.Append(paragraphProperties33);
            paragraph33.Append(proofError35);
            paragraph33.Append(run45);
            paragraph33.Append(proofError36);
            paragraph33.Append(run46);
            paragraph33.Append(proofError37);
            paragraph33.Append(run47);
            paragraph33.Append(proofError38);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "158B6BD6", TextId = "77777777" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines31 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize81 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties34.Append(runFonts79);
            paragraphMarkRunProperties34.Append(fontSize81);

            paragraphProperties34.Append(spacingBetweenLines31);
            paragraphProperties34.Append(paragraphMarkRunProperties34);
            ProofError proofError39 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run48 = new Run();

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts80 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize82 = new FontSize() { Val = "28" };

            runProperties48.Append(runFonts80);
            runProperties48.Append(fontSize82);
            Text text48 = new Text();
            text48.Text = "ของกิจการในการเปลี่ยนทรัพย์สินที่มีอยู่ไปเป็นเงินสดเพื่อแสดง";

            run48.Append(runProperties48);
            run48.Append(text48);
            ProofError proofError40 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph34.Append(paragraphProperties34);
            paragraph34.Append(proofError39);
            paragraph34.Append(run48);
            paragraph34.Append(proofError40);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "20B4A498", TextId = "77777777" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines32 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize83 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties35.Append(runFonts81);
            paragraphMarkRunProperties35.Append(fontSize83);

            paragraphProperties35.Append(spacingBetweenLines32);
            paragraphProperties35.Append(paragraphMarkRunProperties35);
            ProofError proofError41 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run49 = new Run();

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts82 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize84 = new FontSize() { Val = "28" };

            runProperties49.Append(runFonts82);
            runProperties49.Append(fontSize84);
            Text text49 = new Text();
            text49.Text = "ถึงความสามารถในการชำระภาระผูกพัน";

            run49.Append(runProperties49);
            run49.Append(text49);
            ProofError proofError42 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run50 = new Run();

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts83 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize85 = new FontSize() { Val = "28" };

            runProperties50.Append(runFonts83);
            runProperties50.Append(fontSize85);
            Text text50 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text50.Text = " (";

            run50.Append(runProperties50);
            run50.Append(text50);
            ProofError proofError43 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run51 = new Run();

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts84 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize86 = new FontSize() { Val = "28" };

            runProperties51.Append(runFonts84);
            runProperties51.Append(fontSize86);
            Text text51 = new Text();
            text51.Text = "หนี้";

            run51.Append(runProperties51);
            run51.Append(text51);
            ProofError proofError44 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run52 = new Run();

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts85 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize87 = new FontSize() { Val = "28" };

            runProperties52.Append(runFonts85);
            runProperties52.Append(fontSize87);
            Text text52 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text52.Text = ") ";

            run52.Append(runProperties52);
            run52.Append(text52);
            ProofError proofError45 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run53 = new Run();

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize88 = new FontSize() { Val = "28" };

            runProperties53.Append(runFonts86);
            runProperties53.Append(fontSize88);
            Text text53 = new Text();
            text53.Text = "ระยะสั้นของกิจการ";

            run53.Append(runProperties53);
            run53.Append(text53);
            ProofError proofError46 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(proofError41);
            paragraph35.Append(run49);
            paragraph35.Append(proofError42);
            paragraph35.Append(run50);
            paragraph35.Append(proofError43);
            paragraph35.Append(run51);
            paragraph35.Append(proofError44);
            paragraph35.Append(run52);
            paragraph35.Append(proofError45);
            paragraph35.Append(run53);
            paragraph35.Append(proofError46);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "4D1B295E", TextId = "77777777" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines33 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize89 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties36.Append(runFonts87);
            paragraphMarkRunProperties36.Append(fontSize89);

            paragraphProperties36.Append(spacingBetweenLines33);
            paragraphProperties36.Append(paragraphMarkRunProperties36);
            ProofError proofError47 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run54 = new Run();

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize90 = new FontSize() { Val = "28" };

            runProperties54.Append(runFonts88);
            runProperties54.Append(fontSize90);
            Text text54 = new Text();
            text54.Text = "ในอนาคต";

            run54.Append(runProperties54);
            run54.Append(text54);
            ProofError proofError48 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(proofError47);
            paragraph36.Append(run54);
            paragraph36.Append(proofError48);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph33);
            tableCell17.Append(paragraph34);
            tableCell17.Append(paragraph35);
            tableCell17.Append(paragraph36);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders18 = new TableCellBorders();
            TopBorder topBorder19 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder19 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder19 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder19 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders18.Append(topBorder19);
            tableCellBorders18.Append(leftBorder19);
            tableCellBorders18.Append(bottomBorder19);
            tableCellBorders18.Append(rightBorder19);
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin18 = new TableCellMargin();
            TopMargin topMargin18 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin18 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin18 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin18 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin18.Append(topMargin18);
            tableCellMargin18.Append(leftMargin18);
            tableCellMargin18.Append(bottomMargin18);
            tableCellMargin18.Append(rightMargin18);
            HideMark hideMark18 = new HideMark();

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellBorders18);
            tableCellProperties18.Append(shading18);
            tableCellProperties18.Append(tableCellMargin18);
            tableCellProperties18.Append(hideMark18);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "260EE64C", TextId = "77777777" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines34 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize91 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties37.Append(runFonts89);
            paragraphMarkRunProperties37.Append(fontSize91);

            paragraphProperties37.Append(spacingBetweenLines34);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run55 = new Run();

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts90 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize92 = new FontSize() { Val = "28" };

            runProperties55.Append(runFonts90);
            runProperties55.Append(fontSize92);
            Text text55 = new Text();
            text55.Text = "- Liquidity ratio";

            run55.Append(runProperties55);
            run55.Append(text55);

            paragraph37.Append(paragraphProperties37);
            paragraph37.Append(run55);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "24A5C1B6", TextId = "77777777" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines35 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts91 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize93 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties38.Append(runFonts91);
            paragraphMarkRunProperties38.Append(fontSize93);

            paragraphProperties38.Append(spacingBetweenLines35);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            Run run56 = new Run();

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts92 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize94 = new FontSize() { Val = "28" };

            runProperties56.Append(runFonts92);
            runProperties56.Append(fontSize94);
            Text text56 = new Text();
            text56.Text = "- Change in TCA";

            run56.Append(runProperties56);
            run56.Append(text56);

            paragraph38.Append(paragraphProperties38);
            paragraph38.Append(run56);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "06051D44", TextId = "77777777" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines36 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize95 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties39.Append(runFonts93);
            paragraphMarkRunProperties39.Append(fontSize95);

            paragraphProperties39.Append(spacingBetweenLines36);
            paragraphProperties39.Append(paragraphMarkRunProperties39);

            Run run57 = new Run();

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts94 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize96 = new FontSize() { Val = "28" };

            runProperties57.Append(runFonts94);
            runProperties57.Append(fontSize96);
            Text text57 = new Text();
            text57.Text = "- Investment asset per policyholder liability";

            run57.Append(runProperties57);
            run57.Append(text57);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run57);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "0D406B62", TextId = "77777777" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines37 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize97 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties40.Append(runFonts95);
            paragraphMarkRunProperties40.Append(fontSize97);

            paragraphProperties40.Append(spacingBetweenLines37);
            paragraphProperties40.Append(paragraphMarkRunProperties40);

            Run run58 = new Run();

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize98 = new FontSize() { Val = "28" };

            runProperties58.Append(runFonts96);
            runProperties58.Append(fontSize98);
            Text text58 = new Text();
            text58.Text = "- Bad debt per total income";

            run58.Append(runProperties58);
            run58.Append(text58);

            paragraph40.Append(paragraphProperties40);
            paragraph40.Append(run58);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph37);
            tableCell18.Append(paragraph38);
            tableCell18.Append(paragraph39);
            tableCell18.Append(paragraph40);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "1900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders19 = new TableCellBorders();
            TopBorder topBorder20 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder20 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder20 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder20 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders19.Append(topBorder20);
            tableCellBorders19.Append(leftBorder20);
            tableCellBorders19.Append(bottomBorder20);
            tableCellBorders19.Append(rightBorder20);
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin19 = new TableCellMargin();
            TopMargin topMargin19 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin19 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin19 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin19 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin19.Append(topMargin19);
            tableCellMargin19.Append(leftMargin19);
            tableCellMargin19.Append(bottomMargin19);
            tableCellMargin19.Append(rightMargin19);
            HideMark hideMark19 = new HideMark();

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellBorders19);
            tableCellProperties19.Append(shading19);
            tableCellProperties19.Append(tableCellMargin19);
            tableCellProperties19.Append(hideMark19);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "7B64527C", TextId = "77777777" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines38 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts runFonts97 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize99 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties41.Append(runFonts97);
            paragraphMarkRunProperties41.Append(fontSize99);

            paragraphProperties41.Append(spacingBetweenLines38);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            Run run59 = new Run();

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts98 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize100 = new FontSize() { Val = "28" };

            runProperties59.Append(runFonts98);
            runProperties59.Append(fontSize100);
            Text text59 = new Text();
            text59.Text = "- Investment asset per reserve";

            run59.Append(runProperties59);
            run59.Append(text59);

            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run59);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "52BD4E9D", TextId = "77777777" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines39 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize101 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties42.Append(runFonts99);
            paragraphMarkRunProperties42.Append(fontSize101);

            paragraphProperties42.Append(spacingBetweenLines39);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            Run run60 = new Run();

            RunProperties runProperties60 = new RunProperties();
            RunFonts runFonts100 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize102 = new FontSize() { Val = "28" };

            runProperties60.Append(runFonts100);
            runProperties60.Append(fontSize102);
            Text text60 = new Text();
            text60.Text = "- Surrender ratio";

            run60.Append(runProperties60);
            run60.Append(text60);

            paragraph42.Append(paragraphProperties42);
            paragraph42.Append(run60);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph41);
            tableCell19.Append(paragraph42);

            tableRow5.Append(tableCell16);
            tableRow5.Append(tableCell17);
            tableRow5.Append(tableCell18);
            tableRow5.Append(tableCell19);

            TableRow tableRow6 = new TableRow() { RsidTableRowAddition = "00561265", ParagraphId = "7CB02C7B", TextId = "77777777" };

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "1500", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders20 = new TableCellBorders();
            TopBorder topBorder21 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder21 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder21 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder21 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders20.Append(topBorder21);
            tableCellBorders20.Append(leftBorder21);
            tableCellBorders20.Append(bottomBorder21);
            tableCellBorders20.Append(rightBorder21);
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin20 = new TableCellMargin();
            TopMargin topMargin20 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin20 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin20 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin20 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin20.Append(topMargin20);
            tableCellMargin20.Append(leftMargin20);
            tableCellMargin20.Append(bottomMargin20);
            tableCellMargin20.Append(rightMargin20);
            HideMark hideMark20 = new HideMark();

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(tableCellBorders20);
            tableCellProperties20.Append(shading20);
            tableCellProperties20.Append(tableCellMargin20);
            tableCellProperties20.Append(hideMark20);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "4B89A244", TextId = "77777777" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines40 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            RunFonts runFonts101 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize103 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties43.Append(runFonts101);
            paragraphMarkRunProperties43.Append(fontSize103);

            paragraphProperties43.Append(spacingBetweenLines40);
            paragraphProperties43.Append(paragraphMarkRunProperties43);

            Run run61 = new Run();

            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize104 = new FontSize() { Val = "28" };

            runProperties61.Append(runFonts102);
            runProperties61.Append(fontSize104);
            Text text61 = new Text();
            text61.Text = "4. Reinsurance";

            run61.Append(runProperties61);
            run61.Append(text61);

            paragraph43.Append(paragraphProperties43);
            paragraph43.Append(run61);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph43);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "4800", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders21 = new TableCellBorders();
            TopBorder topBorder22 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder22 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder22 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder22 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders21.Append(topBorder22);
            tableCellBorders21.Append(leftBorder22);
            tableCellBorders21.Append(bottomBorder22);
            tableCellBorders21.Append(rightBorder22);
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin21 = new TableCellMargin();
            TopMargin topMargin21 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin21 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin21 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin21 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin21.Append(topMargin21);
            tableCellMargin21.Append(leftMargin21);
            tableCellMargin21.Append(bottomMargin21);
            tableCellMargin21.Append(rightMargin21);
            HideMark hideMark21 = new HideMark();

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(tableCellBorders21);
            tableCellProperties21.Append(shading21);
            tableCellProperties21.Append(tableCellMargin21);
            tableCellProperties21.Append(hideMark21);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "7A22F4B9", TextId = "77777777" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines41 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            RunFonts runFonts103 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize105 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties44.Append(runFonts103);
            paragraphMarkRunProperties44.Append(fontSize105);

            paragraphProperties44.Append(spacingBetweenLines41);
            paragraphProperties44.Append(paragraphMarkRunProperties44);
            ProofError proofError49 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run62 = new Run();

            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize106 = new FontSize() { Val = "28" };

            runProperties62.Append(runFonts104);
            runProperties62.Append(fontSize106);
            Text text62 = new Text();
            text62.Text = "การวิเคราะห์สัดส่วนการประกันภัยต่อและการกระจุกตัวของบริษัท";

            run62.Append(runProperties62);
            run62.Append(text62);
            ProofError proofError50 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph44.Append(paragraphProperties44);
            paragraph44.Append(proofError49);
            paragraph44.Append(run62);
            paragraph44.Append(proofError50);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "2A135E3C", TextId = "77777777" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines42 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize107 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties45.Append(runFonts105);
            paragraphMarkRunProperties45.Append(fontSize107);

            paragraphProperties45.Append(spacingBetweenLines42);
            paragraphProperties45.Append(paragraphMarkRunProperties45);
            ProofError proofError51 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run63 = new Run();

            RunProperties runProperties63 = new RunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize108 = new FontSize() { Val = "28" };

            runProperties63.Append(runFonts106);
            runProperties63.Append(fontSize108);
            Text text63 = new Text();
            text63.Text = "ประกันภัยต่อรวมถึงความสามารถในเร่งรัดจัดเก็บเงินค้างรับจาก";

            run63.Append(runProperties63);
            run63.Append(text63);
            ProofError proofError52 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph45.Append(paragraphProperties45);
            paragraph45.Append(proofError51);
            paragraph45.Append(run63);
            paragraph45.Append(proofError52);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "4965CB41", TextId = "77777777" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines43 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            RunFonts runFonts107 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize109 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties46.Append(runFonts107);
            paragraphMarkRunProperties46.Append(fontSize109);

            paragraphProperties46.Append(spacingBetweenLines43);
            paragraphProperties46.Append(paragraphMarkRunProperties46);
            ProofError proofError53 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run64 = new Run();

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts108 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize110 = new FontSize() { Val = "28" };

            runProperties64.Append(runFonts108);
            runProperties64.Append(fontSize110);
            Text text64 = new Text();
            text64.Text = "การประกันภัยต่อ";

            run64.Append(runProperties64);
            run64.Append(text64);

            Run run65 = new Run() { RsidRunAddition = "00B57ECF" };

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize111 = new FontSize() { Val = "28" };

            runProperties65.Append(runFonts109);
            runProperties65.Append(fontSize111);
            Text text65 = new Text();
            text65.Text = "wwwwwwwwwwwwwwwwwwwwwwww";

            run65.Append(runProperties65);
            run65.Append(text65);
            ProofError proofError54 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph46.Append(paragraphProperties46);
            paragraph46.Append(proofError53);
            paragraph46.Append(run64);
            paragraph46.Append(run65);
            paragraph46.Append(proofError54);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphAddition = "00B57ECF", RsidRunAdditionDefault = "00B57ECF", ParagraphId = "532AAA21", TextId = "3847FE5C" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines44 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            RunFonts runFonts110 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize112 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties47.Append(runFonts110);
            paragraphMarkRunProperties47.Append(fontSize112);

            paragraphProperties47.Append(spacingBetweenLines44);
            paragraphProperties47.Append(paragraphMarkRunProperties47);
            ProofError proofError55 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run66 = new Run();

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts111 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize113 = new FontSize() { Val = "28" };

            runProperties66.Append(runFonts111);
            runProperties66.Append(fontSize113);
            Text text66 = new Text();
            text66.Text = "Ddddddddddddddddddddddddddddddddddddddd";

            run66.Append(runProperties66);
            run66.Append(text66);
            ProofError proofError56 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph47.Append(paragraphProperties47);
            paragraph47.Append(proofError55);
            paragraph47.Append(run66);
            paragraph47.Append(proofError56);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphAddition = "00B57ECF", RsidRunAdditionDefault = "00B57ECF", ParagraphId = "76291248", TextId = "7F3336F6" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines45 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            RunFonts runFonts112 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize114 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties48.Append(runFonts112);
            paragraphMarkRunProperties48.Append(fontSize114);

            paragraphProperties48.Append(spacingBetweenLines45);
            paragraphProperties48.Append(paragraphMarkRunProperties48);
            ProofError proofError57 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run67 = new Run();

            RunProperties runProperties67 = new RunProperties();
            RunFonts runFonts113 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize115 = new FontSize() { Val = "28" };

            runProperties67.Append(runFonts113);
            runProperties67.Append(fontSize115);
            Text text67 = new Text();
            text67.Text = "Wwwwwwwwwwwwwwwwwwwwww";

            run67.Append(runProperties67);
            run67.Append(text67);
            ProofError proofError58 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph48.Append(paragraphProperties48);
            paragraph48.Append(proofError57);
            paragraph48.Append(run67);
            paragraph48.Append(proofError58);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphAddition = "00B57ECF", RsidRunAdditionDefault = "00B57ECF", ParagraphId = "61C31C4F", TextId = "51B9ED6A" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines46 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            RunFonts runFonts114 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize116 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties49.Append(runFonts114);
            paragraphMarkRunProperties49.Append(fontSize116);

            paragraphProperties49.Append(spacingBetweenLines46);
            paragraphProperties49.Append(paragraphMarkRunProperties49);
            ProofError proofError59 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run68 = new Run();

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts115 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize117 = new FontSize() { Val = "28" };

            runProperties68.Append(runFonts115);
            runProperties68.Append(fontSize117);
            Text text68 = new Text();
            text68.Text = "Wwwwwwwwwwwwwwwwwwwwwwwwwwwww";

            run68.Append(runProperties68);
            run68.Append(text68);
            ProofError proofError60 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph49.Append(paragraphProperties49);
            paragraph49.Append(proofError59);
            paragraph49.Append(run68);
            paragraph49.Append(proofError60);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphAddition = "00B57ECF", RsidRunAdditionDefault = "00B57ECF", ParagraphId = "57E68A05", TextId = "2628FA3C" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines47 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            RunFonts runFonts116 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize118 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties50.Append(runFonts116);
            paragraphMarkRunProperties50.Append(fontSize118);

            paragraphProperties50.Append(spacingBetweenLines47);
            paragraphProperties50.Append(paragraphMarkRunProperties50);
            ProofError proofError61 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run69 = new Run();

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts117 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize119 = new FontSize() { Val = "28" };

            runProperties69.Append(runFonts117);
            runProperties69.Append(fontSize119);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text69 = new Text();
            text69.Text = "Wwwwwwwwwwwwwwwwwwwwww";

            run69.Append(runProperties69);
            run69.Append(lastRenderedPageBreak1);
            run69.Append(text69);
            ProofError proofError62 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph50.Append(paragraphProperties50);
            paragraph50.Append(proofError61);
            paragraph50.Append(run69);
            paragraph50.Append(proofError62);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphAddition = "00B57ECF", RsidRunAdditionDefault = "00B57ECF", ParagraphId = "4412F58B", TextId = "0DBFF5B7" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines48 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            RunFonts runFonts118 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize120 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties51.Append(runFonts118);
            paragraphMarkRunProperties51.Append(fontSize120);

            paragraphProperties51.Append(spacingBetweenLines48);
            paragraphProperties51.Append(paragraphMarkRunProperties51);
            ProofError proofError63 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run70 = new Run();

            RunProperties runProperties70 = new RunProperties();
            RunFonts runFonts119 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize121 = new FontSize() { Val = "28" };

            runProperties70.Append(runFonts119);
            runProperties70.Append(fontSize121);
            Text text70 = new Text();
            text70.Text = "wwwwwwwwwwwwwwwwwwwwwww";

            run70.Append(runProperties70);
            run70.Append(text70);
            ProofError proofError64 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            paragraph51.Append(paragraphProperties51);
            paragraph51.Append(proofError63);
            paragraph51.Append(run70);
            paragraph51.Append(proofError64);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph44);
            tableCell21.Append(paragraph45);
            tableCell21.Append(paragraph46);
            tableCell21.Append(paragraph47);
            tableCell21.Append(paragraph48);
            tableCell21.Append(paragraph49);
            tableCell21.Append(paragraph50);
            tableCell21.Append(paragraph51);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders22 = new TableCellBorders();
            TopBorder topBorder23 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder23 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder23 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder23 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders22.Append(topBorder23);
            tableCellBorders22.Append(leftBorder23);
            tableCellBorders22.Append(bottomBorder23);
            tableCellBorders22.Append(rightBorder23);
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin22 = new TableCellMargin();
            TopMargin topMargin22 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin22 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin22 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin22 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin22.Append(topMargin22);
            tableCellMargin22.Append(leftMargin22);
            tableCellMargin22.Append(bottomMargin22);
            tableCellMargin22.Append(rightMargin22);
            HideMark hideMark22 = new HideMark();

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(tableCellBorders22);
            tableCellProperties22.Append(shading22);
            tableCellProperties22.Append(tableCellMargin22);
            tableCellProperties22.Append(hideMark22);

            Paragraph paragraph52 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "2662B5DB", TextId = "77777777" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines49 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            RunFonts runFonts120 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize122 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties52.Append(runFonts120);
            paragraphMarkRunProperties52.Append(fontSize122);

            paragraphProperties52.Append(spacingBetweenLines49);
            paragraphProperties52.Append(paragraphMarkRunProperties52);

            Run run71 = new Run();

            RunProperties runProperties71 = new RunProperties();
            RunFonts runFonts121 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize123 = new FontSize() { Val = "28" };

            runProperties71.Append(runFonts121);
            runProperties71.Append(fontSize123);
            LastRenderedPageBreak lastRenderedPageBreak2 = new LastRenderedPageBreak();
            Text text71 = new Text();
            text71.Text = "- Reinsurance premium receivable ratio";

            run71.Append(runProperties71);
            run71.Append(lastRenderedPageBreak2);
            run71.Append(text71);

            paragraph52.Append(paragraphProperties52);
            paragraph52.Append(run71);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "33E75CAA", TextId = "77777777" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines50 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            RunFonts runFonts122 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize124 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties53.Append(runFonts122);
            paragraphMarkRunProperties53.Append(fontSize124);

            paragraphProperties53.Append(spacingBetweenLines50);
            paragraphProperties53.Append(paragraphMarkRunProperties53);

            Run run72 = new Run();

            RunProperties runProperties72 = new RunProperties();
            RunFonts runFonts123 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize125 = new FontSize() { Val = "28" };

            runProperties72.Append(runFonts123);
            runProperties72.Append(fontSize125);
            Text text72 = new Text();
            text72.Text = "- Reinsurance income ratio";

            run72.Append(runProperties72);
            run72.Append(text72);

            paragraph53.Append(paragraphProperties53);
            paragraph53.Append(run72);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "4DDAB1E2", TextId = "77777777" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines51 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            RunFonts runFonts124 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize126 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties54.Append(runFonts124);
            paragraphMarkRunProperties54.Append(fontSize126);

            paragraphProperties54.Append(spacingBetweenLines51);
            paragraphProperties54.Append(paragraphMarkRunProperties54);

            Run run73 = new Run();

            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts125 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize127 = new FontSize() { Val = "28" };

            runProperties73.Append(runFonts125);
            runProperties73.Append(fontSize127);
            LastRenderedPageBreak lastRenderedPageBreak3 = new LastRenderedPageBreak();
            Text text73 = new Text();
            text73.Text = "- Change in loss ratio after reinsurance";

            run73.Append(runProperties73);
            run73.Append(lastRenderedPageBreak3);
            run73.Append(text73);

            paragraph54.Append(paragraphProperties54);
            paragraph54.Append(run73);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph52);
            tableCell22.Append(paragraph53);
            tableCell22.Append(paragraph54);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "1900", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders23 = new TableCellBorders();
            TopBorder topBorder24 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder24 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder24 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder24 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableCellBorders23.Append(topBorder24);
            tableCellBorders23.Append(leftBorder24);
            tableCellBorders23.Append(bottomBorder24);
            tableCellBorders23.Append(rightBorder24);
            Shading shading23 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            TableCellMargin tableCellMargin23 = new TableCellMargin();
            TopMargin topMargin23 = new TopMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            LeftMargin leftMargin23 = new LeftMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin23 = new BottomMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };
            RightMargin rightMargin23 = new RightMargin() { Width = "48", Type = TableWidthUnitValues.Dxa };

            tableCellMargin23.Append(topMargin23);
            tableCellMargin23.Append(leftMargin23);
            tableCellMargin23.Append(bottomMargin23);
            tableCellMargin23.Append(rightMargin23);
            HideMark hideMark23 = new HideMark();

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellBorders23);
            tableCellProperties23.Append(shading23);
            tableCellProperties23.Append(tableCellMargin23);
            tableCellProperties23.Append(hideMark23);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "59E719EB", TextId = "77777777" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines52 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            RunFonts runFonts126 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize128 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties55.Append(runFonts126);
            paragraphMarkRunProperties55.Append(fontSize128);

            paragraphProperties55.Append(spacingBetweenLines52);
            paragraphProperties55.Append(paragraphMarkRunProperties55);

            Run run74 = new Run();

            RunProperties runProperties74 = new RunProperties();
            RunFonts runFonts127 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize129 = new FontSize() { Val = "28" };

            runProperties74.Append(runFonts127);
            runProperties74.Append(fontSize129);
            LastRenderedPageBreak lastRenderedPageBreak4 = new LastRenderedPageBreak();
            Text text74 = new Text();
            text74.Text = "- Retention ratio";

            run74.Append(runProperties74);
            run74.Append(lastRenderedPageBreak4);
            run74.Append(text74);

            paragraph55.Append(paragraphProperties55);
            paragraph55.Append(run74);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphAddition = "00561265", RsidRunAdditionDefault = "00981400", ParagraphId = "11DA3F04", TextId = "77777777" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines53 = new SpacingBetweenLines() { After = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            RunFonts runFonts128 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize130 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties56.Append(runFonts128);
            paragraphMarkRunProperties56.Append(fontSize130);

            paragraphProperties56.Append(spacingBetweenLines53);
            paragraphProperties56.Append(paragraphMarkRunProperties56);

            Run run75 = new Run();

            RunProperties runProperties75 = new RunProperties();
            RunFonts runFonts129 = new RunFonts() { Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize131 = new FontSize() { Val = "28" };

            runProperties75.Append(runFonts129);
            runProperties75.Append(fontSize131);
            Text text75 = new Text();
            text75.Text = "- Change in loss ratio after reinsurance";

            run75.Append(runProperties75);
            run75.Append(text75);

            paragraph56.Append(paragraphProperties56);
            paragraph56.Append(run75);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph55);
            tableCell23.Append(paragraph56);

            tableRow6.Append(tableCell20);
            tableRow6.Append(tableCell21);
            tableRow6.Append(tableCell22);
            tableRow6.Append(tableCell23);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);
            table1.Append(tableRow6);
            Paragraph paragraph57 = new Paragraph() { RsidParagraphAddition = "00981400", RsidRunAdditionDefault = "00981400", ParagraphId = "0DA75516", TextId = "77777777" };

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "00981400" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1080U, Bottom = 1440, Left = (UInt32Value)1080U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(table1);
            body1.Append(paragraph57);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
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
            Zoom zoom1 = new Zoom() { Percent = "50" };
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
            Rsid rsid1 = new Rsid() { Val = "003C4E3D" };
            Rsid rsid2 = new Rsid() { Val = "003E1962" };
            Rsid rsid3 = new Rsid() { Val = "003F2413" };
            Rsid rsid4 = new Rsid() { Val = "0047664C" };
            Rsid rsid5 = new Rsid() { Val = "004D5E7B" };
            Rsid rsid6 = new Rsid() { Val = "00561265" };
            Rsid rsid7 = new Rsid() { Val = "00671749" };
            Rsid rsid8 = new Rsid() { Val = "00745E8C" };
            Rsid rsid9 = new Rsid() { Val = "00981400" };
            Rsid rsid10 = new Rsid() { Val = "00AD54DD" };
            Rsid rsid11 = new Rsid() { Val = "00B57ECF" };
            Rsid rsid12 = new Rsid() { Val = "00F56D5B" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);
            rsids1.Append(rsid12);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin24 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin24 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin24);
            mathProperties1.Append(rightMargin24);
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
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "2F0F16B8" };
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
            settings1.Append(documentId1);
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
            RunFonts runFonts130 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize132 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "ja-JP", Bidi = "th-TH" };

            runPropertiesBaseStyle1.Append(runFonts130);
            runPropertiesBaseStyle1.Append(fontSize132);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines54 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines54);

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
            TopMargin topMargin24 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin24 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin24);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin24);
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

        // Generates content of numberingDefinitionsPart1.
        private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
        {
            Numbering numbering1 = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            numbering1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            numbering1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            numbering1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            numbering1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            numbering1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            numbering1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            numbering1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            numbering1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            numbering1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            numbering1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            numbering1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            numbering1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            numbering1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            numbering1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            numbering1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            numbering1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            numbering1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            numbering1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            numbering1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            numbering1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            numbering1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            numbering1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            numbering1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            numbering1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid1 = new Nsid() { Val = "5CCD5ED0" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "BDBA2F62" };

            Level level1 = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation2 = new Indentation() { Start = "720", Hanging = "360" };

            previousParagraphProperties1.Append(indentation2);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts131 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize133 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties1.Append(runFonts131);
            numberingSymbolRunProperties1.Append(fontSize133);
            numberingSymbolRunProperties1.Append(fontSizeComplexScript2);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1 };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText2 = new LevelText() { Val = "%1.%2." };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation3 = new Indentation() { Start = "1440", Hanging = "360" };

            previousParagraphProperties2.Append(indentation3);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts132 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize134 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties2.Append(runFonts132);
            numberingSymbolRunProperties2.Append(fontSize134);
            numberingSymbolRunProperties2.Append(fontSizeComplexScript3);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties2);

            Level level3 = new Level() { LevelIndex = 2, Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText3 = new LevelText() { Val = "%1.%2.%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation4 = new Indentation() { Start = "2160", Hanging = "180" };

            previousParagraphProperties3.Append(indentation4);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts133 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize135 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties3.Append(runFonts133);
            numberingSymbolRunProperties3.Append(fontSize135);
            numberingSymbolRunProperties3.Append(fontSizeComplexScript4);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);
            level3.Append(numberingSymbolRunProperties3);

            Level level4 = new Level() { LevelIndex = 3, Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
            Indentation indentation5 = new Indentation() { Start = "2880", Hanging = "360" };

            previousParagraphProperties4.Append(indentation5);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts134 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize136 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties4.Append(runFonts134);
            numberingSymbolRunProperties4.Append(fontSize136);
            numberingSymbolRunProperties4.Append(fontSizeComplexScript5);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);
            level4.Append(numberingSymbolRunProperties4);

            Level level5 = new Level() { LevelIndex = 4, Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5." };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
            Indentation indentation6 = new Indentation() { Start = "3600", Hanging = "360" };

            previousParagraphProperties5.Append(indentation6);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts135 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize137 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties5.Append(runFonts135);
            numberingSymbolRunProperties5.Append(fontSize137);
            numberingSymbolRunProperties5.Append(fontSizeComplexScript6);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);
            level5.Append(numberingSymbolRunProperties5);

            Level level6 = new Level() { LevelIndex = 5, Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
            Indentation indentation7 = new Indentation() { Start = "4320", Hanging = "180" };

            previousParagraphProperties6.Append(indentation7);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts136 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize138 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties6.Append(runFonts136);
            numberingSymbolRunProperties6.Append(fontSize138);
            numberingSymbolRunProperties6.Append(fontSizeComplexScript7);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);
            level6.Append(numberingSymbolRunProperties6);

            Level level7 = new Level() { LevelIndex = 6, Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
            Indentation indentation8 = new Indentation() { Start = "5040", Hanging = "360" };

            previousParagraphProperties7.Append(indentation8);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts137 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize139 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties7.Append(runFonts137);
            numberingSymbolRunProperties7.Append(fontSize139);
            numberingSymbolRunProperties7.Append(fontSizeComplexScript8);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);
            level7.Append(numberingSymbolRunProperties7);

            Level level8 = new Level() { LevelIndex = 7, Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8." };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
            Indentation indentation9 = new Indentation() { Start = "5760", Hanging = "360" };

            previousParagraphProperties8.Append(indentation9);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts138 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize140 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties8.Append(runFonts138);
            numberingSymbolRunProperties8.Append(fontSize140);
            numberingSymbolRunProperties8.Append(fontSizeComplexScript9);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);
            level8.Append(numberingSymbolRunProperties8);

            Level level9 = new Level() { LevelIndex = 8, Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
            Indentation indentation10 = new Indentation() { Start = "6480", Hanging = "180" };

            previousParagraphProperties9.Append(indentation10);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts139 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "TH SarabunPSK", HighAnsi = "TH SarabunPSK", ComplexScript = "TH SarabunPSK" };
            FontSize fontSize141 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

            numberingSymbolRunProperties9.Append(runFonts139);
            numberingSymbolRunProperties9.Append(fontSize141);
            numberingSymbolRunProperties9.Append(fontSizeComplexScript10);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);
            level9.Append(numberingSymbolRunProperties9);

            abstractNum1.Append(nsid1);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);
            abstractNum1.Append(level1);
            abstractNum1.Append(level2);
            abstractNum1.Append(level3);
            abstractNum1.Append(level4);
            abstractNum1.Append(level5);
            abstractNum1.Append(level6);
            abstractNum1.Append(level7);
            abstractNum1.Append(level8);
            abstractNum1.Append(level9);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 0 };

            numberingInstance1.Append(abstractNumId1);

            numbering1.Append(abstractNum1);
            numbering1.Append(numberingInstance1);

            numberingDefinitionsPart1.Numbering = numbering1;
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

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

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

            Font font1 = new Font() { Name = "TH SarabunPSK" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020B0500040200020003" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "A100006F", UnicodeSignature1 = "5000205A", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00010183", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Yu Mincho" };
            AltName altName1 = new AltName() { Val = "游明朝" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "02020400000000000000" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "800002E7", UnicodeSignature1 = "2AC7FCFF", UnicodeSignature2 = "00000012", UnicodeSignature3 = "00000000", CodePageSignature0 = "0002009F", CodePageSignature1 = "00000000" };

            font4.Append(altName1);
            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Cordia New" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020B0304020202020204" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "81000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00010001", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Yu Gothic Light" };
            AltName altName2 = new AltName() { Val = "游ゴシック Light" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020B0300000000000000" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "2AC7FDFF", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "0002009F", CodePageSignature1 = "00000000" };

            font6.Append(altName2);
            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            Font font7 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number7 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily7 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch7 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature7 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C000247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font7.Append(panose1Number7);
            font7.Append(fontCharSet7);
            font7.Append(fontFamily7);
            font7.Append(pitch7);
            font7.Append(fontSignature7);

            Font font8 = new Font() { Name = "Angsana New" };
            Panose1Number panose1Number8 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily8 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch8 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature8 = new FontSignature() { UnicodeSignature0 = "81000003", UnicodeSignature1 = "00000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "00010001", CodePageSignature1 = "00000000" };

            font8.Append(panose1Number8);
            font8.Append(fontCharSet8);
            font8.Append(fontFamily8);
            font8.Append(pitch8);
            font8.Append(fontSignature8);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);

            fontTablePart1.Fonts = fonts1;
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


    }
}
