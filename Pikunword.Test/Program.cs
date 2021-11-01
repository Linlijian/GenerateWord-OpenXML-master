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
                numbering_type = new string[] {  },
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
            var pmp = new List<PikunTableCellProperties>();
            var t = new List<PikunTableGrid>();

            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "MERGE4",
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                grid_span = 2,
                justification = Pikun.justificationRight,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentBottom
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "C1",
                prop = new string[] { Pikun.paragraphBold }                
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "C2",
                prop = new string[] { Pikun.paragraphBold }
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "MERGE",
                prop = new string[] { Pikun.paragraphBold },
                grid_span = 2
            });           
            t.Add(new PikunTableGrid
            {
                rId = 12,
                grid_column = "3740",
                table_cell_width = "500",
                table_cell_properties = pmp
            });

            pmp = new List<PikunTableCellProperties>();
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                vertical_merge_child = true,
                grid_span = 2
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "R3",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "R4",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "R5",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "R6",
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                grid_column = "1882",
                table_cell_width = "500",
                table_cell_properties = pmp
            });

            pmp = new List<PikunTableCellProperties>();
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S!",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S@",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S#",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S$",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S^",
            });
            pmp.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S&",
                fill = Pikun.Red_Dirt
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                grid_column = "1882",
                table_cell_width = "500",
                table_cell_properties = pmp
            });

            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 8,
                    table_style = "TableGrid",
                    table_width = "1870",
                    have_table_cell_margin = false,
                    table_grid = t
                }
            });
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
