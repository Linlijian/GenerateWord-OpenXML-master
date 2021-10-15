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
            GenerateMainDocumentPartContent(mainDocumentPart _dto);
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
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1, pikunword_dto _dto)
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

            Paragraph paragraph1 = newLineImage("00000001", "00000002", "rId4", imagePart1Data);

            Paragraph paragraph2 = newLine("00000001", "00000002");
            Paragraph paragraph5 = newLine("00000001", "00000002");
            Paragraph paragraph6 = newLine("00000001", "00000002");
            Paragraph paragraph3 = newLine("00000001", "00000001", "mmCGチャンネル animetic");
            Paragraph paragraph4 = newLine("00000001", "00000001", "สำนักงานคณะกรรมการกำกับและส่งเสริมการประกอบธุรกิจประกันภัย (สำนักงาน คปภ.) มีบทบาทหน้าที่ในการส่งเสริมสนับสนุนให้ธุรกิจประกันภัยมีบทบาทสร้างเสริมความแข็งแกร่งให้ระบบเศรษฐกิจ สังคมของประเทศและคุณภาพชีวิตที่ดีของประชาชน รวมทั้งผลักดันให้ธุรกิจประกันภัยก้าวหน้า"); //ซ้ำได้มีเพื่อ??????

            Paragraph paragraph7 = newLineImage("00000001", "00000002", "rId7", imagePart2Data, Help.wrapSquare, 14, 1, 0, 0);

            Paragraph paragraph8 = newLine("00000001", "00000001", "mmCGチャンネル animetic", Help.paragraphBold);
            Paragraph paragraph9 = newLine("00000001", "00000001", "mmCGチャンネル animetic", Help.paragraphItalic);
            Paragraph paragraph10 = newLine("00000001", "00000001", "mmCGチャンネル animetic", Help.paragraphUnderline);

            //ใส่ txt_prop อย่างเดียวใน 1 paragraph
            Paragraph paragraph11 = newLine("00000001", "00000001", "mmCGチャンネル animetic สำนักงานคณะกรรมการกำกับและส่งเสริมการประกอบธุรกิจประกันภัย (สำนักงาน คปภ.)", "animetic", Help.paragraphBold);

            //ใส่ txt_prop หลายๆ อย่างในครั้งเดียว
            System.Collections.Generic.List<string[]> txt_prop = new System.Collections.Generic.List<string[]>();
            txt_prop.Add(new string[] { "การประเมินความเสี่ยง Internal Rating ประกอบด้วย", Help.paragraphNormal });
            txt_prop.Add(new string[] { " การประเมินเชิงปริมาณ 80%", Help.paragraphBold });
            txt_prop.Add(new string[] { " และ ", Help.paragraphNormal });
            txt_prop.Add(new string[] { "การประเมินเชิงคุณภาพ 20%", Help.paragraphUnderline });
            Paragraph paragraph12 = newLine("00000001", "00000001", txt_prop);


            System.Collections.Generic.List<string[]> txt_prop2 = new System.Collections.Generic.List<string[]>();
            txt_prop2.Add(new string[] { "การประเมินความเสี่ยง Internal Rating ประกอบด้วย", Help.paragraphNormal });
            txt_prop2.Add(new string[] { " การประเมินเชิงปริมาณ 80%", Help.paragraphBold, Help.paragraphItalic });
            txt_prop2.Add(new string[] { " และ ", Help.paragraphNormal });
            txt_prop2.Add(new string[] { "การประเมินเชิงคุณภาพ 20%", Help.paragraphUnderline, Help.paragraphBold, Help.paragraphItalic });
            Paragraph paragraph13 = newLineManyprop("00000001", "00000001", txt_prop2);


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
            body1.Append(paragraph7);
            body1.Append(paragraph2);
            body1.Append(paragraph3);
            body1.Append(paragraph6);
            body1.Append(paragraph4);
            body1.Append(paragraph5);
            body1.Append(paragraph8);
            body1.Append(paragraph9);
            body1.Append(paragraph10);
            body1.Append(paragraph11);
            body1.Append(paragraph12);
            body1.Append(paragraph13);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }





        #endregion
    }
}
