using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Generate_Word_Report.dll
{
    class testuse
    {
        static void Main2(string[] args)
        {
            //test qwe = new test();
            //qwe.Model.template1.Text = "";

            pikunword_dto dto = new pikunword_dto();
            dto.Pictures.Add(new PikunPicture
            {
                sizeY = 0
            });

            pikunword da = new pikunword();
            da.word.Pictures.Add(new PikunPicture
            {
                sizeY = 0
            });
            da.Generate(dto);


            

            da.word.Model.path = @"c:\\aaa.worx";

            var models = new List<PikunParagraphManyProp>();
            models.Add(new PikunParagraphManyProp{
                text = "การประเมินความเสี่ยง Internal Rating ประกอบด้วย",
                font = "Itim",
                font_size = 12,
                prop = new string[] { Help.paragraphNormal }
            });
            models.Add(new PikunParagraphManyProp
            {
                text = " การประเมินเชิงปริมาณ 80%",
                font = "Itim",
                justification = Help.justificationLeft,
                font_size = 12,
                prop = new string[] { Help.paragraphItalic, Help.paragraphBold }
            });

            da.word.Models.Add(new pikunword_model
            {
                rId = 1,
                execut_type = pikun_execut_function.newLineManyprop,
                paragraph = new PikunParagraph {
                    rId = 1,
                    many_prop = models
                },
            });
            da.word.Models.Add(new pikunword_model
            {
                rId = 1,
                execut_type = pikun_execut_function.newLineManyprop,
                paragraph = new PikunParagraph
                {
                    rId = 1,
                    text = "Hello"
                },
            });

            da.word.Model.execut_type = pikun_execut_type.create_packet;
            da.Generate(dto);
        }
    }
}
