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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Generate_Word_Report.dll
{
    public class pikunword
    {
        public pikunword_dto word { get; set; }
        private const int rId = 5;
        private const string ExtendedFilePropertiesPart_rId = "rId1";

        public pikunword()
        {
            word = new pikunword_dto();
        }
        public pikunword_dto Generate(pikunword_dto _dto)
        {
            switch (_dto.Model.execut_type)
            {
                case pikun_execut_type.create_packet: return create_packet(_dto);
            }
            return _dto;
        }

        public pikunword_dto xxx(pikunword_dto _dto)
        {
           //do something
            return _dto;
        }

        public pikunword_dto create_packet(pikunword_dto _dto)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(_dto.Model.path, WordprocessingDocumentType.Document))
            {
                create_part(package, _dto);
            }

            //int i = 0;
            //foreach (var m in _dto.Models)
            //{
            //    switch (m.execut_type)
            //    {
            //        case pikun_execut_type.picture: return picture(_dto);
            //        case pikun_execut_type.paragraphs: return paragraphs(_dto);
            //    }
            //}
            return _dto;
        }

        public pikunword_dto picture(pikunword_dto _dto)
        {
            //do something
            return _dto;
        }

        public pikunword_dto paragraphs(pikunword_dto _dto)
        {
           
            return _dto;
        }



        #region method
        
        public void create_part(WordprocessingDocument _document, pikunword_dto _dto)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart = _document.AddNewPart<ExtendedFilePropertiesPart>(ExtendedFilePropertiesPart_rId);
            GenerateExtendedFilePropertiesPartContent(extendedFilePropertiesPart, _dto);

            MainDocumentPart mainDocumentPart = _document.AddMainDocumentPart();
            GenerateMainDocumentPartContent(mainDocumentPart, _dto);

            //image
            //bullet-numbering
        }

        private void GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart, pikunword_dto _dto)
        {
            Ap.Properties properties = new Ap.Properties();
            properties.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            properties.Append(_dto.ExtendedFileProperties.template);
            properties.Append(_dto.ExtendedFileProperties.totalTime);
            properties.Append(_dto.ExtendedFileProperties.pages);
            properties.Append(_dto.ExtendedFileProperties.words);
            properties.Append(_dto.ExtendedFileProperties.characters);
            properties.Append(_dto.ExtendedFileProperties.application);
            properties.Append(_dto.ExtendedFileProperties.documentSecurity);
            properties.Append(_dto.ExtendedFileProperties.lines);
            properties.Append(_dto.ExtendedFileProperties.paragraphs);
            properties.Append(_dto.ExtendedFileProperties.scaleCrop);
            properties.Append(_dto.ExtendedFileProperties.company);
            properties.Append(_dto.ExtendedFileProperties.linksUpToDate);
            properties.Append(_dto.ExtendedFileProperties.charactersWithSpaces);
            properties.Append(_dto.ExtendedFileProperties.sharedDocument);
            properties.Append(_dto.ExtendedFileProperties.hyperlinksChanged);
            properties.Append(_dto.ExtendedFileProperties.applicationVersion);

            extendedFilePropertiesPart.Properties = properties;
        }
        private void GenerateMainDocumentPartContent(MainDocumentPart mainDocumentPart, pikunword_dto _dto)
        {
            #region xml service
            Document document = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            document.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            document.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            document.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            document.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            document.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            document.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            document.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            document.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            document.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            document.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            document.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            document.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            #endregion

            Body body = new Body();
            //===============================================
            //  start body zone
            //===============================================
            foreach (var p in _dto.Models)
            {
                if (p.execut_type == pikun_execut_type.paragraphs) 
                {
                    if(p.paragraph.numbering_id != 0)
                    {
                        if(p.paragraph.many_prop.Count > 0)
                        {
                            Paragraph paragraph = newLineManyProp(p.paragraph);
                            body.Append(paragraph);
                        }
                        else
                        {
                            Paragraph paragraph = newLine(p.paragraph);
                            body.Append(paragraph);
                        }
                        //newLineListParagraph

                        /*
                         * newLine enter space
                         * newLine Normal | all single prop
                         * newLine Manyprop
                         * newLine numbering | all single prop
                         * newLine bullet | all single prop
                         * newLine numbering prop
                         * newLine bullet prop
                         * สร้าง paragraphProperties เป็น fn แล้วส่งกลับมาทีเดียวจะสั้นกว่า
                         */

                    }
                }
            }
            //===============================================
            //  end body zone
            //===============================================

            SectionProperties sectionProperties = new SectionProperties() { RsidR = "003F2413" };
            PageSize pageSize = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U };
            PageMargin pageMargin = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns = new Columns() { Space = "720" };
            DocGrid docGrid = new DocGrid() { LinePitch = 360 };

            sectionProperties.Append(pageSize);
            sectionProperties.Append(pageMargin);
            sectionProperties.Append(columns);
            sectionProperties.Append(docGrid);


            //===============================================
            //  start body append
            //===============================================

            //===============================================
            //  end body append
            //===============================================

            body.Append(sectionProperties);

            document.Append(body);

            mainDocumentPart.Document = document;
        }

        private Paragraph newLineManyProp(PikunParagraph p)
        {
            Paragraph paragraph = new Paragraph() { RsidParagraphMarkRevision = "PIKUNRPM", RsidParagraphAddition = "PIKUNRPA", RsidParagraphProperties = "PIKUNRPP", RsidRunAdditionDefault = "PIKUNRAD", ParagraphId = "PIKUNP" + p.rId, TextId = "PIKUNT" + p.rId };

            RunProperties runProperties;
            ParagraphProperties paragraphProperties;
            ParagraphMarkRunProperties paragraphMarkRunProperties;

            #region ใส่คุณลักษณะมากกว่าหนึ่งอย่าง
            int i = 0;
            foreach (var txt in p.many_prop)
            {
                runProperties = new RunProperties();
                paragraphProperties = new ParagraphProperties();
                paragraphMarkRunProperties = new ParagraphMarkRunProperties();

                for (int tp = 0; tp < txt.prop.Length; tp++)
                {
                    if (txt.prop[tp] == Help.paragraphBold)
                    {
                        Bold bold = new Bold();
                        BoldComplexScript boldComplexScript = new BoldComplexScript();

                        paragraphMarkRunProperties.Append(bold);
                        paragraphMarkRunProperties.Append(boldComplexScript);


                        Bold bold2 = new Bold();
                        BoldComplexScript boldComplexScript2 = new BoldComplexScript();

                        runProperties.Append(bold2);
                        runProperties.Append(boldComplexScript2);
                    }
                    else if (txt.prop[tp] == Help.paragraphItalic)
                    {
                        Italic italic = new Italic();
                        ItalicComplexScript italicComplexScript = new ItalicComplexScript();

                        paragraphMarkRunProperties.Append(italic);
                        paragraphMarkRunProperties.Append(italicComplexScript);


                        Italic italic2 = new Italic();
                        ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();

                        runProperties.Append(italic2);
                        runProperties.Append(italicComplexScript2);
                    }
                    else if (txt.prop[tp] == Help.paragraphUnderline)
                    {
                        Underline underline = new Underline() { Val = UnderlineValues.Single };
                        paragraphMarkRunProperties.Append(underline);

                        Underline underline2 = new Underline() { Val = UnderlineValues.Single };
                        runProperties.Append(underline2);
                    }
                }

                if (!txt.font.IsNullOrEmpty())
                {
                    RunFonts runFonts = new RunFonts() { Ascii = txt.font, HighAnsi = txt.font, ComplexScript = txt.font };
                    paragraphMarkRunProperties.Append(runFonts);

                    RunFonts runFonts2 = new RunFonts() { Ascii = txt.font, HighAnsi = txt.font, ComplexScript = txt.font };
                    runProperties.Append(runFonts2);
                }

                if (!txt.font_size.IsNullOrEmpty())
                {
                    FontSize fontSize = new FontSize() { Val = (txt.font_size * 2).ToString() };
                    FontSizeComplexScript fontSizeComplexScript = new FontSizeComplexScript() { Val = (txt.font_size * 2).ToString() };
                    paragraphMarkRunProperties.Append(fontSize);
                    paragraphMarkRunProperties.Append(fontSizeComplexScript);

                    FontSize fontSize1 = new FontSize() { Val = (txt.font_size * 2).ToString() };
                    FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = (txt.font_size * 2).ToString() };
                    runProperties.Append(fontSize1);
                    runProperties.Append(fontSizeComplexScript1);
                }

                if (!txt.color.IsNullOrEmpty())
                {
                    Color color = new Color() { Val = txt.color };
                    paragraphMarkRunProperties.Append(color);

                    Color color1 = new Color() { Val = txt.color };
                    runProperties.Append(color1);
                }

                if (!txt.justification.IsNullOrEmpty())
                {
                    if (txt.justification == Help.justificationCenter)
                    {
                        Justification justification = new Justification() { Val = JustificationValues.Center };
                        paragraphProperties.Append(justification);
                    }
                    else if (txt.justification == Help.justificationRight)
                    {
                        Justification justification = new Justification() { Val = JustificationValues.Right };
                        paragraphProperties.Append(justification);
                    }
                    else
                    {
                        Justification justification = new Justification() { Val = JustificationValues.Left };
                        paragraphProperties.Append(justification);
                    }
                }

                if (!txt.highlight.IsNullOrEmpty())
                {
                    Highlight highlight;
                    switch (txt.highlight)
                    {
                        case "Black":
                            highlight = new Highlight() { Val = HighlightColorValues.Black };
                            runProperties.Append(highlight);
                            break;
                        case "Blue":
                            highlight = new Highlight() { Val = HighlightColorValues.Blue };
                            runProperties.Append(highlight);
                            break;
                        case "Cyan":
                            highlight = new Highlight() { Val = HighlightColorValues.Cyan };
                            runProperties.Append(highlight);
                            break;
                        case "DarkBlue":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkBlue };
                            runProperties.Append(highlight);
                            break;
                        case "DarkCyan":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkCyan };
                            runProperties.Append(highlight);
                            break;
                        case "DarkGray":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkGray };
                            runProperties.Append(highlight);
                            break;
                        case "DarkGreen":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkGreen };
                            runProperties.Append(highlight);
                            break;
                        case "DarkMagenta":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkMagenta };
                            runProperties.Append(highlight);
                            break;
                        case "DarkRed":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkRed };
                            runProperties.Append(highlight);
                            break;
                        case "DarkYellow":
                            highlight = new Highlight() { Val = HighlightColorValues.DarkYellow };
                            runProperties.Append(highlight);
                            break;
                        case "Green":
                            highlight = new Highlight() { Val = HighlightColorValues.Green };
                            runProperties.Append(highlight);
                            break;
                        case "LightGray":
                            highlight = new Highlight() { Val = HighlightColorValues.LightGray };
                            runProperties.Append(highlight);
                            break;
                        case "Magenta":
                            highlight = new Highlight() { Val = HighlightColorValues.Magenta };
                            runProperties.Append(highlight);
                            break;
                        case "None":
                            highlight = new Highlight() { Val = HighlightColorValues.None };
                            runProperties.Append(highlight);
                            break;
                        case "Red":
                            highlight = new Highlight() { Val = HighlightColorValues.Red };
                            runProperties.Append(highlight);
                            break;
                        case "White":
                            highlight = new Highlight() { Val = HighlightColorValues.White };
                            runProperties.Append(highlight);
                            break;
                        case "Yellow":
                            highlight = new Highlight() { Val = HighlightColorValues.Yellow };
                            runProperties.Append(highlight);
                            break;
                        default:
                            break;
                    }
                }

                paragraphProperties.Append(paragraphMarkRunProperties);

                Run run = new Run() { RsidRunProperties = "PIKUNRRP" };

                Text text;
                if (i == 0)
                {
                    text = new Text();
                }
                else
                {
                    text = new Text() { Space = SpaceProcessingModeValues.Preserve };
                }

                text.Text = txt.text;

                if (txt.prop.First() != Help.paragraphNormal)
                {
                    run.Append(runProperties);
                }

                run.Append(text);

                ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };

                if (i == 0)
                {
                    paragraph.Append(paragraphProperties);
                    paragraph.Append(proofError1);
                }

                paragraph.Append(run);

                i++;
            }
            #endregion
           
            return paragraph;
        }
        private Paragraph newLine(PikunParagraph p)
        {
            return new Paragraph();
        }

        #endregion
    }
}
