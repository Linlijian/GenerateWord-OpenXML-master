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

            #region setting numbering
            //=============================================================================================
            pk.word.NumberingDefinitions.Add(new PikunNumberingDefinitions
            {
                numbering_type = new string[] { },
                number_format_values = Pikun.numberFormatValuesDecimal,
                font = "TH SarabunPSK",
                font_size = 14
            });
            //=============================================================================================
            #endregion

            #region paragrap 1
            //=============================================================================================
            models.Add(new PikunParagraphManyProp
            {
                text = "สรุปรายงานวิเคราะห์",
                font_size = 14,
                font = "TH SarabunPSK",
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineManyprop,
                paragraph = new PikunParagraph
                {
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
                    font = "TH SarabunPSK",
                    text = "            รายงานการวิเคราะห์ จะประกอบด้วยรายละเอียดจำนวน 8 ส่วน ดังต่อไปนี้",
                    prop = Pikun.paragraphBold
                }
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
                    font = "TH SarabunPSK",
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
                    font = "TH SarabunPSK",
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
                    font = "TH SarabunPSK",
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
                font = "TH SarabunPSK",
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " การประเมินเชิงปริมาณ 80%",
                font_size = 14,
                font = "TH SarabunPSK",
                prop = new string[] { Pikun.paragraphBold }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " และ",
                font = "TH SarabunPSK",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " การประเมินเชิงคุณภาพ 20%",
                font = "TH SarabunPSK",
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
                    font = "TH SarabunPSK",
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            //=============================================================================================
            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "การประเมินเชิงปริมาณ",
                font = "TH SarabunPSK",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " มีการพิจารณา 5 องค์ประกอบหลักดังตางรางด้านล่าง โดยมีการให้",
                font = "TH SarabunPSK",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " น้ำหนักในแต่ละองค์ประกอบสำหรับการประเมินบริษัทประกันวินาศภัยและบริษัทประกันชีวิตที่แตกต่างกัน",
                font = "TH SarabunPSK",
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
            #endregion

            #region table 1
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
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                spacing_between_lines = true,
                vertical_merge = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "ค่าน้ำหนักของวินาศภัย",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "ค่าน้ำหนักของชีวิต",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "ปี 2562",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                vertical_merge = true,
                grid_span = 2,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c2.Add(new PikunTableCellProperties
            {
                rId = 17,
                vertical_merge_child = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 18,
                vertical_merge_child = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 19,
                vertical_merge_child = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 19,
                text = "คะแนน (เต็ม 5)*",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 19,
                multi_line = true,
                texts = new string[] { "ผลรวม" ,"(คะแนน x น้ำหนัก)" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "1. Profitability",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "25%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "30%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "2. Capital adequacy",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "30%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "25%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "3. Liquidity",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "20%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "15%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "4. Reinsurance",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "15%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "5%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "5. Investment",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "10%",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "25%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });


            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รวม",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "100%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "100%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c8.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c9.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "คะแนนหลังคิดค่าน้ำหนักการประเมินเชิงปริมาณ (ร้อยละ 80)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                grid_span = 4,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c9.Add(new PikunTableCellProperties
            {
                rId = 13,
                font_size = 14,
                text = "x.xx (.80)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                spacing_between_lines = true,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            #endregion

            #region add c
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
            #endregion

            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 8,
                    table_cell_width_auto = false,
                    table_style = "TableGrid",
                    have_table_cell_margin = false,
                    grid_column_size = 5,
                    grid_column = new string[] { "3000", "1000", "1000", "2000", "2000" }, //เป็น 0 เพราะต้องการ defualt 1870
                    table_grid = t
                }
            });
            //=============================================================================================
            #endregion

            #region paragrap 2
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 2,
                    font_size = 14,
                    one_text_mamy_paragrap = true, // ใส่ ' ' ต่อท้ายจะดีมาก
                    font = "TH SarabunPSK",
                    texts = new string[] { "                * คะแนนระดับ 5 หมายถึง บริษัทมีระดับความเสี่ยงสูงมากเมื่อเปรียบเทียบกับบริษัทอื่นในอุตสาหกรรม ซึ่งคะแนนที่ได้จะ ",
                        "คำนวณมาจากระบบ SIIRA โดยแต่ละรายการจะมีองค์ประกอบย่อย เช่น อัตราส่วนหรือตัวชี้วัดต่าง ๆ เป็นต้น ระบบจะนำค่าของ ",
                        "องค์ประกอบย่อยแต่ละรายการไปเปรียบเทียบกับค่าเฉลี่ยอุตสาหกรรม เพื่อคำนวณออกมาเป็นคะแนน ซึ่งคณะทำงานของสายวิเคราะห์ " ,
                        "ธุรกิจประกันภัยเป็นผู้กำหนดระดับคะแนนในแต่ละช่วง" }
                }
            });
            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "การประเมินเชิงคุณภาพ",
                font = "TH SarabunPSK",
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " มีการพิจารณาคุณภาพด้านการบริหารจัดการและกระบวนการควบคุม",
                font = "TH SarabunPSK",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = "ของบริษัท 4 มุมมอง ตามตารางด้านล่างโดยให้น้ำหนักในแต่ละมุมมองสำหรับการประเมินบริษัทประกันวินาศภัย",
                font = "TH SarabunPSK",
                font_size = 14,
                prop = new string[] { Pikun.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " และบริษัทประกันชีวิตเหมือนกัน",
                font = "TH SarabunPSK",
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
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 2
                }
            });
            //=============================================================================================
            #endregion

            #region table 2
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
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "ค่าน้ำหนัก",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 15,
                multi_line = true,
                texts = new string[] { "คะแนน", "(เต็ม 5 คะแนน) *" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 16,
                multi_line = true,
                texts = new string[] { "คะแนน", "(คะแนน x น้ำหนัก)" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
            });

            c2.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "1. ความเพียงพอของการบริหารจัดการองค์กร (Adequacy of corporate management)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "25%",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "2. ความเพียงพอของระบบการควบคุมภายในและการตรวจสอบภายใน (Adequacy of Internal control system andInternal Audit Activities)",
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "25%",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c4.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "3. ความเพียงพอของระบบการบริหารความเสี่ยง (Adequacy of risk management system)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "25%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c4.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c5.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "4. External audit and Actuary risks",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                font_size = 14
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "25%",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            c5.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });

            c6.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รวม",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                font_size = 14,
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 14,
                text = "100%",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 15,
                text = "",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c6.Add(new PikunTableCellProperties
            {
                rId = 16,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "คะแนนหลังคิดค่าน้ำหนักการประเมินเชิงคุณภาพ (ร้อยละ 20)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                grid_span = 3,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c7.Add(new PikunTableCellProperties
            {
                rId = 13,
                font_size = 14,
                text = "x.xx (.20)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            #endregion

            #region row
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
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c3,
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c4,
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c5,
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c6,
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c7,
            });
            #endregion

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
                    grid_column = new string[] { "5800", "1100", "1500", "1680" },
                }
            });
            //=============================================================================================
            #endregion

            #region paaragrap 3
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    font_size = 14,
                    font = "TH SarabunPSK",
                    text = "                * คะแนนระดับ 5 หมายถึง บริษัทมีระดับความเสี่ยงสูงมากเมื่อเปรียบเทียบกับเกณฑ์พื้นฐานที่คณะทำงาน ของสายวิเคราะห์ ธุรกิจประกันภัยกำหนดไว้เป็น Checklists ให้ผู้วิเคราะห์ทำการประเมินโดยพิจารณาจากข้อมูลใน รายงานที่บริษัทนำส่ง เช่น รายงานการ บริหารความเสี่ยงแบบองค์รวม (ORSA) รวมถึงข้อมูลที่ได้รับจากสายตรวจสอบ ที่ทำการเข้าตรวจ ณ ที่ทำการบริษัท เป็นต้น"
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 22
                }
            });

            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "คะแนนการประเมินรวม",
                font = "TH SarabunPSK",
                font_size = 16,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineManyprop,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    many_prop = models,
                }
            });
            //=============================================================================================
            #endregion

            #region table 3
            //=============================================================================================
            c1 = new List<PikunTableCellProperties>();
            c2 = new List<PikunTableCellProperties>();
            c3 = new List<PikunTableCellProperties>();
            t = new List<PikunTableGrid>();

            #region c
            c1.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "การประเมินเชิงปริมาณ (80%)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c1.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });

            c2.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "การประเมินเชิงคุณภาพ (20%)",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c2.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });

            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "คะแนนรวม",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            c3.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "x.xx",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter,
                spacing_between_lines = true
            });
            #endregion

            #region t
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
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = c3,
            });
            #endregion

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
                    grid_column_size = 2,
                    grid_column = new string[] { "7000", "3000" },
                }
            });
            //=============================================================================================
            #endregion

            #region paragrap 4
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 22
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "สรุปผลการประเมินความเสี่ยง",
                    font = "TH SarabunPSK",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    numbering_level_reference = 0
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    text = "สรุปผลการประเมินความเสี่ยง",
                    prop = Pikun.paragraphUnderline,
                    font_size = 14,
                    font = "TH SarabunPSK"
                }
            });
            //=============================================================================================
            #endregion

            #region image 1
            //=============================================================================================
            string path = AppDomain.CurrentDomain.BaseDirectory + "image1.png";
            string imagePart2Data = Pikun.BitmapToBase64String(path);
            pk.word.Pictures.Add(new PikunPicture
            {
                rId = 7,
                base64image = imagePart2Data
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineImageNoFormat,
                picture = new PikunPicture
                {
                    rId = 7, // << ต้องเหมือนกันกับ setting
                    base64image = imagePart2Data
                    //layout_option = Help.wrapSquare,
                    //horizontal_position = 14,
                    //vertical_position = 0,
                    //sizeX = 0,
                    //sizeY = 0
                }
            });
            //=============================================================================================
            #endregion

            #region paragrap 5
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    font_size = 14,
                    font = "TH SarabunPSK",
                    text = "           นำคะแนนการประเมินรวมจากผลการประเมินเชิงปริมาณและเชิงคุณภาพข้างต้น มาสรุปเป็นผลการประเมินความเสี่ยง และวิเคราะห์องค์ประกอบของการประเมินเชิงปริมาณและเชิงคุณภาพในแต่ละรายการ ที่มีผลต่อการประเมินความเสี่ยงของบริษัท ดังนี้"
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                }
            });
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "ความเสี่ยงที่วัดจากตัวชี้วัดเชิงปริมาณ",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    font = "TH SarabunPSK",
                    numbering_level_reference = 1
                }
            });
            //=============================================================================================
            #endregion

            #region table 4
            //=============================================================================================
            var row = new List<PikunTableCellProperties>();
            t = new List<PikunTableGrid>();

            #region r1
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "องค์ประกอบหลัก",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                vertical_merge = true,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รายละเอียดการวิเคราะห์",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                vertical_merge = true,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "องค์ประกอบย่อย",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                grid_span = 2,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r2
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                vertical_merge_child = true
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                vertical_merge_child = true
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "บริษัทประกันวินาศภัย",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "บริษัทประกันชีวิต",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                prop = new string[] { Pikun.paragraphBold },
                spacing_between_lines = true,
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r3
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "1. Profitability",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                texts = new string[] { "การวิเคราะห์พอร์ตการรับประกันภัยและผลการดำเนินงาน ",
                    "(เบี้ยประกันภัยค่าสินไหมทดแทนและค่าใช้จ่ายของบริษัท) รวมถึง ",
                    "ปัจจัยความเสี่ยงที่เกี่ยวข้องกับการรับประกันภัย อาทิ ช่องทาง",
                    "การขายและการติดตามเบี้ยประกันภัยค้างรับ เป็นต้น"},
                multi_line = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Loss ratio", "- Expense ratio", "- Premium receivable before impairment", "- Return on equity (ROE)" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Expense ratio", "- Change in net written premium", "- Change in single premium", "- Return on equity (ROE)" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r4
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "2. Capital adequacy",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                texts = new string[] { "การวิเคราะห์องค์ประกอบของเงินกองทุนที่สามารถนำมาใช้ได้ทั้งหมด ",
                    "(TCA) และเงินกองทุนที่ต้องดำรงทั้งหมด (TCR) ซึ่งต้องสามารถ",
                    "แสดงให้เห็นถึงปัจจัยเสี่ยงที่กระทบต่อระดับเงินกองทุนของบริษัท"},
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                multi_line = true,
                font = "TH SarabunPSK",
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Capital adequacy ratio (CAR)", "- Change in TCA", "- Net written premium per TCA", "- Commission income per TCA" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Capital adequacy ratio (CAR)", "- Change in TCA" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r5
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "3. Liquidity",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                texts = new string[] { "การวิเคราะห์กระแสเงินสดรับ-จ่ายของบริษัท และวัดความสามารถ",
                    "ของกิจการในการเปลี่ยนทรัพย์สินที่มีอยู่ไปเป็นเงินสดเพื่อแสดง",
                    "ถึงความสามารถในการชำระภาระผูกพัน (หนี้) ระยะสั้นของกิจการ",
                    "ในอนาคต"},
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                multi_line = true,
                font = "TH SarabunPSK",
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Liquidity ratio", "- Change in TCA", "- Investment asset per policyholder liability", "- Bad debt per total income" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Investment asset per reserve", "- Surrender ratio" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });

            #endregion

            #region r6
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "4. Reinsurance",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                texts = new string[] { "การวิเคราะห์สัดส่วนการประกันภัยต่อและการกระจุกตัวของบริษัท",
                    "ประกันภัยต่อรวมถึงความสามารถในเร่งรัดจัดเก็บเงินค้างรับจาก",
                    "การประกันภัยต่อ"},
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                multi_line = true,
                font = "TH SarabunPSK",
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Reinsurance premium receivable ratio", "- Reinsurance income ratio", "- Change in loss ratio after reinsurance" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Retention ratio", "- Change in loss ratio after reinsurance" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r7
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "5. Investment",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                texts = new string[] { "การวิเคราะห์พอร์ตการลงทุนและผลตอบแทนที่ได้จากการลงทุนรวมถึง",
                    "ความเสี่ยงจากการลงทุนของบริษัท อาทิ ความผันผวนของสินทรัพย์",
                    "ลงทุนสินทรัพย์ลงทุนที่มีระดับความน่าเชื่อค่อนข้างต่ำและการกระจุกตัว",
                    "ในสินทรัพย์ลงทุน เป็นต้น"},
                multi_line = true,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Return on investment (ROI)", "- Return and profit on investment" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "- Duration gap", "- Return on investment (ROI)", "- Return and profit on" },
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 28,
                    table_style = "TableGrid",
                    have_table_cell_margin = false,
                    table_cell_width_auto = false,
                    table_grid = t,
                    grid_column_size = 4,
                    grid_column = new string[] { "1500", "4800", "2000", "1900" },
                }
            });
            //=============================================================================================
            #endregion

            #region paragrap 6
            //=============================================================================================
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "ความเสี่ยงที่วัดจากตัวชี้วัดในเชิงคุณภาพ",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    font = "TH SarabunPSK",
                    numbering_level_reference = 1
                }
            });
            //=============================================================================================
            #endregion

            #region table 5
            row = new List<PikunTableCellProperties>();
            t = new List<PikunTableGrid>();

            #region r1
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "มุมมอง",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "รายละเอียดการวิเคราะห์",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r2
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "1) ความเพียงพอของการบริหารจัดการองค์กร (Adequacy of corporate management)",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop 
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "พิจารณาโครงสร้างและองค์ประกอบต่างๆ ของบริษัท แผนธุรกิจ และช่องทางการขาย เป็นต้น",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r3
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "2) ความเพียงพอของระบบการควบคุมภายในและการตรวจสอบ",
                    "ภายใน(Adequacy of Internalcontrol system and Internal Audit Activities)" },
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "พิจารณาองค์ประกอบคณะกรรมการตรวจสอบ และคุณภาพของการนำส่งรายงานทางการเงิน",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r4
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "3) ความเพียงพอของระบบการบริหารความเสี่ยง (Adequacy of risk management system)",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "พิจารณาจากผลการตรวจสอบของบริษัท นโยบายบริหารความเสี่ยง และผลทดสอบสภาวะวิกฤต(Stress Test) เป็นต้น",
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font_size = 14,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r5
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "4) External audit and Actuary risks",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "วิเคราะห์คุณภาพและแสดงความเห็นของผู้สอบบัญชี คุณภาพนักคณิตศาสตร์ประกันภัย เป็นต้น",
                font = "TH SarabunPSK",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentTop
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 28,
                    table_style = "TableGrid",
                    have_table_cell_margin = false,
                    table_cell_width_auto = false,
                    table_grid = t,
                    grid_column_size = 2,
                    grid_column = new string[] { "5000", "5000" },
                }
            });

            #endregion

            #region paragrap 7
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    font_size = 14,
                    font = "TH SarabunPSK",
                    text = "           ทั้งนี้ รายละเอียดการวิเคราะห์องค์ประกอบของการประเมินเชิงปริมาณและเชิงคุณภาพในแต่ละรายการจะขึ้นอยู่กับดุลยพินิจ ของผู้วิเคราะห์แต่ละคนในการพิจารณาเลือกประเด็นวิเคราะห์ที่เห็นว่าสำคัญและมีความเสี่ยง"
                }
            });

            var space = new pikunword_model
            {
                execut_type = pikun_execut_function.newLine,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                }
            };
            pk.word.Models.Add(space);

            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "ตารางความเสี่ยงของบริษัทที่กระทบต่ออุตสาหกรรม",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    font = "TH SarabunPSK",
                    numbering_level_reference = 0
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    font_size = 14,
                    font = "TH SarabunPSK",
                    text = "           จากสรุปผลคะแนนความเสี่ยง นำมาทำ Risk Matrix เปรียบเทียบ Rating ของบริษัทกับผลกระทบในระดับอุตสาหกรรม"
                }
            });

            pk.word.Models.Add(space);

            models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp
            {
                text = "ตารางความเสี่ยงของบริษัทที่กระทบต่ออุตสาหกรรม : Risk Mapping",
                font = "TH SarabunPSK",
                font_size = 16,
                prop = new string[] { Pikun.paragraphBold, Pikun.paragraphUnderline }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineManyprop,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    many_prop = models,
                }
            });
            #endregion

            #region table 6
            row = new List<PikunTableCellProperties>();
            t = new List<PikunTableGrid>();

            #region r1
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Size",
                font_size = 14,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Probability/Effect",
                font_size = 14,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] {"Not important","(1)" },
                font_size = 14,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "Slightly important", "(2)" },
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "Important", "(3)" },
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "Serious", "(4)" },
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                multi_line = true,
                texts = new string[] { "Critical", "(5)" },
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r2
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "L",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Very High",
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low med",
                fill = Pikun.Green_Thumb,
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Medium",
                fill = Pikun.Yellow_W3C,
                font_size = 14,
                font = "TH SarabunPSK",
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Med Hi",
                fill = Pikun.Orange_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "High",
                fill = Pikun.Red_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "High",
                fill = Pikun.Red_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r3
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "M++",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "High",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low",
                fill = Pikun.Green_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low med",
                fill = Pikun.Green_Thumb,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Medium",
                fill = Pikun.Yellow_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Med Hi",
                fill = Pikun.Orange_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "High",
                fill = Pikun.Red_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r4
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "M",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Possible",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low",
                fill = Pikun.Green_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low med",
                fill = Pikun.Green_Thumb,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Medium",
                fill = Pikun.Yellow_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Med Hi",
                fill = Pikun.Orange_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Med Hi",
                fill = Pikun.Orange_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r5
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S++",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low Possibility",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low",
                fill = Pikun.Green_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low med",
                fill = Pikun.Green_Thumb,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low med",
                fill = Pikun.Green_Thumb,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Medium",
                fill = Pikun.Yellow_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Med Hi",
                fill = Pikun.Orange_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            #region r6
            row = new List<PikunTableCellProperties>();
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "S",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Very Low",
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationLeft,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low",
                fill = Pikun.Green_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low",
                fill = Pikun.Green_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Low med",
                fill = Pikun.Green_Thumb,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });            
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Medium",
                fill = Pikun.Yellow_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            row.Add(new PikunTableCellProperties
            {
                rId = 13,
                text = "Medium",
                fill = Pikun.Yellow_W3C,
                font_size = 14,
                top_border_size = 6,
                bottom_border_size = 6,
                left_border_size = 6,
                right_border_size = 6,
                font = "TH SarabunPSK",
                spacing_between_lines = true,
                prop = new string[] { Pikun.paragraphBold },
                justification = Pikun.justificationCenter,
                table_cell_vertical_alignment = Pikun.tableCellVerticalAlignmentCenter
            });
            t.Add(new PikunTableGrid
            {
                rId = 12,
                table_cell_properties = row
            });
            #endregion

            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newTable,
                table = new PikunTable
                {
                    rId = 28,
                    table_style = "TableGrid",
                    have_table_cell_margin = false,
                    table_cell_width_auto = false,
                    table_grid = t,
                    grid_column_size = 7,
                    grid_column = new string[] { "900", "1800", "1600", "1800", "1300", "1300", "1300" },
                }
            });
            #endregion

            #region paragrap 8
            pk.word.Models.Add(space);
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "แนวทางกำกับและติดตาม",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    font = "TH SarabunPSK",
                    numbering_level_reference = 0
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    font_size = 14,
                    font = "TH SarabunPSK",
                    text = "           นำเสนอแนวทางในการกำกับและติดตาม ในประเด็นที่พบจากการวิเคราะห์แต่ละตัวชี้วัด"
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNumbering,
                paragraph = new PikunParagraph
                {
                    text = "ข้อคิดเห็นและข้อเสนอแนะ",
                    rId = 5,
                    font_size = 14,
                    numbering_id = 1,
                    font = "TH SarabunPSK",
                    numbering_level_reference = 0
                }
            });
            pk.word.Models.Add(new pikunword_model
            {
                execut_type = pikun_execut_function.newLineNormal,
                paragraph = new PikunParagraph
                {
                    rId = 22,
                    font_size = 14,
                    font = "TH SarabunPSK",
                    text = "           เสนอความคิดเห็นและข้อเสนอแนะอื่นๆ นอกจากนี้ จะมีFact sheet ของบริษัทที่ข้อมูลในภาพรวม โดยสร้างขึ้นมาอัตโนมัติ จากระบบ CRR แนบเป็นเอกสารประกอบ"
                }
            });
            #endregion
            pk.word.Model.execut_type = pikun_execut_type.create_packet;
            pk.word.Model.page_margin = Pikun.pageMarginModerate;
            pk.word.Model.path = AppDomain.CurrentDomain.BaseDirectory + "pikun_test.docx";
            pk.Generate(pk.word);
            //=============================================================================================

        }
    }
}
