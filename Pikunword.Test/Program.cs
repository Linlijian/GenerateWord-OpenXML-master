using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pikunword;

namespace Pikunword.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            pikunword pk = new pikunword();
            var models = new List<PikunParagraphManyProp>();

            //=============================================================================================
            models.Add(new PikunParagraphManyProp
            {
                text = "สรุปรายงานวิเคราะห์",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineManyprop,
                paragraph = new PikunParagraph {
                    rId = 1,
                    many_prop = models
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 2,
                    font_size = 14,
                    text = "            รายงานการวิเคราะห์ จะประกอบด้วยรายละเอียดจำนวน 8 ส่วน ดังต่อไปนี้",
                    prop = Pikun.paragraphBold
                }
            });
            //=============================================================================================
            pk.word.NumberingDefinitions.Add(new PikunNumberingDefinitions
            {
                numbering_type = new string[] { "%1.", "%1.%2.", "%1.%2.%3.", "%1.%2.%3.%4.", "%1.%2.%3.%4.%5.", "%1.%2.%3.%4.%5.%6.", "%1.%2.%3.%4.%5.%6.%7.", "%1.%2.%3.%4.%5.%6.%7.%8.", "%1.%2.%3.%4.%5.%6.%7.%8.%9." },
                number_format_values = Pikun.numberFormatValuesDecimal                
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "Executive Summary สรุปผลการวิเคราะห์ในภาพรวม",
                    rId = 3,
                    font_size = 14,
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "จุดแข็ง / ข้อได้เปรียบ (Strengths) และจุดอ่อน (Weaknesses)",
                    rId = 4,
                    font_size = 14,
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "ข้อมูลบริษัท อาทิ ผู้ถือหุ้น กรรมการ และประเด็นที่ต้องติดตามจากครั้งก่อน เป็นต้น",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            //=============================================================================================
            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "การประเมินความเสี่ยง Internal Rating ประกอบด้วย",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " การประเมินเชิงปริมาณ 80%",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " และ",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " การประเมินเชิงคุณภาพ 20%",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumberingProp,
                paragraph = new PikunParagraph
                {
                    rId = 6,
                    many_prop = models,
                    font_size = 14,
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            //=============================================================================================
            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "การประเมินเชิงปริมาณ",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " มีการพิจารณา 5 องค์ประกอบหลักดังตางรางด้านล่าง โดยมีการให้",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " น้ำหนักในแต่ละองค์ประกอบสำหรับการประเมินบริษัทประกันวินาศภัยและบริษัทประกันชีวิตที่แตกต่างกัน",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumberingProp,
                paragraph = new PikunParagraph
                {
                    rId = 7,
                    many_prop = models,
                    font_size = 14,
                    numbering_id = 1,
                    numbering_level_reference = 1
                }
            });
            //=============================================================================================
            //=============================================================================================
            //=============================================================================================
            //=============================================================================================
            //=============================================================================================
            //=============================================================================================
            //=============================================================================================
            //=============================================================================================
            pk.word.Model.execut_type = pikun_execut_type.create_packet;
            pk.word.Model.path = AppDomain.CurrentDomain.BaseDirectory + "pikun_test.docx";

            pk.Generate(pk.word);
            //=============================================================================================

        }
    }
}
