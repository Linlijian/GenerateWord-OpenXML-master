using Generate_Word_Report.NewGen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using mw = Microsoft.Office.Interop.Word;

namespace Generate_Word_Report
{
    class Program
    {
        static void Main(string[] args)
        {
            custom_one_pic Empty = new custom_one_pic();
            Empty.CreatePackage("custom_one_pic.docx");

            new_line Empty2 = new new_line();
            Empty2.CreatePackage(@"F:\OIC.research\GenerateWord-OpenXML-master\GenerateWord-OpenXML-master\Generate Word Report\new_line.docx");





            var models = new List<dll.PikunParagraphManyProp>();
            models.Add(new dll.PikunParagraphManyProp
            {
                text = "การประเมินความเสี่ยง Internal Rating ประกอบด้วย",
                font = "Itim",
                font_size = 12,
                prop = new string[] { Help.paragraphNormal }
            });
            models.Add(new dll.PikunParagraphManyProp
            {
                text = " การประเมินเชิงปริมาณ 80%",
                font = "Itim",
                justification = Help.justificationLeft,
                font_size = 12,
                prop = new string[] { Help.paragraphItalic, Help.paragraphBold }
            });

            dll.pikunword pk = new dll.pikunword();
            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineNormal,
                paragraph = new dll.PikunParagraph { text = "I LOVE MUK!", rId = 1, prop = Help.paragraphBold, font_size = 20, highlight = Help.highlightColorDarkBlue}
            });
            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineManyprop,
                paragraph = new dll.PikunParagraph {
                    many_prop = models
                }
            });

            /*
             * setting numbering
             */
            pk.word.NumberingDefinitions.Add(new dll.PikunNumberingDefinitions {
                numbering_type = new string[] { },
                number_format_values = Help.numberFormatValuesDecimal,
                font = "Itim"
            });
            pk.word.NumberingDefinitions.Add(new dll.PikunNumberingDefinitions
            {
                numbering_type = new string[] { },
                number_format_values = Help.numberFormatValuesBullet,
                font = "Javanese Text"
            });

            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineNumbering,
                paragraph = new dll.PikunParagraph { text = "I LOVE MUK! Numbering",
                    rId = 1,
                    prop = Help.paragraphBold,
                    font_size = 20,
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineNumbering,
                paragraph = new dll.PikunParagraph { text = "I LOVE MUK! Numbering 2",
                    rId = 1,
                    prop = Help.paragraphBold,
                    font_size = 20,
                    numbering_id = 1,
                    numbering_level_reference = 1
                }
            });
            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineNormal,
                paragraph = new dll.PikunParagraph { text = "I LOVE MUK! B", rId = 1, prop = Help.paragraphBold, font_size = 20, highlight = Help.highlightColorDarkBlue }
            });
            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineNumbering,
                paragraph = new dll.PikunParagraph
                {
                    text = "I LOVE MUK! Numbering B",
                    rId = 1,
                    prop = Help.paragraphBold,
                    font_size = 20,
                    numbering_id = 2,
                    numbering_level_reference = 0
                }
            });
            pk.word.Models.Add(new dll.pikunword_model
            {
                execut_type = dll.pikun_execut_function.newLineNumbering,
                paragraph = new dll.PikunParagraph
                {
                    text = "I LOVE MUK! Numbering 2 B",
                    rId = 1,
                    prop = Help.paragraphBold,
                    font_size = 20,
                    numbering_id = 2,
                    numbering_level_reference = 0
                }
            });
            pk.word.Model.execut_type = dll.pikun_execut_type.create_packet;
            pk.word.Model.path = @"F:\OIC.research\GenerateWord-OpenXML-master\GenerateWord-OpenXML-master\Generate Word Report\bin\Debug\pikun.docx";

            pk.Generate(pk.word);
        }
    }
}
