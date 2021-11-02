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
            var c1 = new List<PikunTableCellProperties>();
            var c2 = new List<PikunTableCellProperties>();
            var c3 = new List<PikunTableCellProperties>();
            var c4 = new List<PikunTableCellProperties>();
            var c5 = new List<PikunTableCellProperties>();
            var c6 = new List<PikunTableCellProperties>();
            var c7 = new List<PikunTableCellProperties>();
            var c8 = new List<PikunTableCellProperties>();
            var c9 = new List<PikunTableCellProperties>();
            var t = new List<PikunTableGrid>();

            #region cell
            c1.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รายการ",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "ค่าน้ำหนักของวินาศภัย",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "ค่าน้ำหนักของชีวิต",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "ปี 2562",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                grid_span = 2,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c2.Add(new PikunTableCellProperties
            {
                rId = 17,
                vertical_merge_child = true,
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 18,
                vertical_merge_child = true,
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 19,
                vertical_merge_child = true,
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 19,
                text = "คะแนน (เต็ม 5)*",
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 19,
                text = "ผลรวม(คะแนน x น้ำหนัก)",
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "1. Profitability",
                font_size = 12
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "25%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "30%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "2. Capital adequacy",
                font_size = 12
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "30%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "25%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "3. Liquidity",
                font_size = 12
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "20%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "15%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "4. Reinsurance",
                font_size = 12
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "15%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "5%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "5. Investment",
                font_size = 12
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "10%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "25%",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });


            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รวม",
                font_size = 12,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "100%",
                font_size = 12,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "100%",
                font_size = 12,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "",
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font_size = 12,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c9.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "คะแนนหลังคิดค่าน้ำหนักการประเมินเชิงปริมาณ (ร้อยละ 80)",
                font_size = 12,
                grid_span = 4,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c9.Add(new PikunTableCellProperties
            {
                rId = 13,
                font_size = 12,
                text = "x.xx (.80)",
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            #endregion

            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c1,
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c2
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c3
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c4
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c5
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c6
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c7
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c8
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c9
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 8,
                    table_cell_width_auto = true,
                    table_style = "TableGrid",
                    have_table_cell_margin = false,
                    grid_column_size = 5,
                    grid_column = new string[] {"0", "0", "0", "0", "0" }, //เป็น 0 เพราะต้องการ defualt 1870
                    table_grid = t
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
                    text = "                * คะแนนระดับ 5 หมายถึง บริษัทมีระดับความเสี่ยงสูงมากเมื่อเปรียบเทียบกับบริษัทอื่นในอุตสาหกรรม ซึ่งคะแนนที่ได้จะคำนวณมาจากระบบ SIIRA โดยแต่ละรายการจะมีองค์ประกอบย่อย เช่น อัตราส่วนหรือตัวชี้วัดต่าง ๆ เป็นต้น ระบบจะนำค่าขององค์ประกอบย่อยแต่ละรายการไปเปรียบเทียบกับค่าเฉลี่ยอุตสาหกรรม เพื่อคำนวณออกมาเป็นคะแนน ซึ่งคณะทำงานของสายวิเคราะห์ธุรกิจประกันภัยเป็นผู้กำหนดระดับคะแนนในแต่ละช่วง"
                }
            });
            //=============================================================================================
            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "การประเมินเชิงคุณภาพ",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " มีการพิจารณาคุณภาพด้านการบริหารจัดการและกระบวนการควบคุม",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = "ของบริษัท 4 มุมมอง ตามตารางด้านล่างโดยให้น้ำหนักในแต่ละมุมมองสำหรับการประเมินบริษัทประกันวินาศภัย",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = "และบริษัทประกันชีวิตเหมือนกัน",
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
            c1 = new List<PikunTableCellProperties>();
            c2 = new List<PikunTableCellProperties>();
            c3 = new List<PikunTableCellProperties>();
            c4 = new List<PikunTableCellProperties>();
            c5 = new List<PikunTableCellProperties>();
            c6 = new List<PikunTableCellProperties>();
            c7 = new List<PikunTableCellProperties>();
            c8 = new List<PikunTableCellProperties>();
            c9 = new List<PikunTableCellProperties>();
            t = new List<PikunTableGrid>();

            #region cell
            c1.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รายการ",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "ค่าน้ำหนัก",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 15,
                multi_line = true,
                texts = new string[] { "คะแนน", "(เต็ม 5 คะแนน) *" },
                font_size = 10,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 16,
                multi_line = true,
                texts = new string[] { "คะแนน", "(คะแนน x น้ำหนัก)" },
                font_size = 10,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });

            c2.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "1. ความเพียงพอของการบริหารจัดการองค์กร (Adequacy of corporate management)",
                font_size = 14
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "25%",
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "x.xx",
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "x.xx",
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            #endregion

            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c1,
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c2,
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 8,
                    table_style = "TableGrid",
                    have_table_cell_margin = false,
                    table_cell_width_auto = false,
                    table_grid = t,
                    grid_column_size = 4,
                    grid_column = new string[] { "5800", "900", "1300", "1380" },
                }
            });
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
