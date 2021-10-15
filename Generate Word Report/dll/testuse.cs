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
            dto.Pictures.Add(new Picture
            {
                sizeY = 0
            });

            pikunword da = new pikunword();
            da.word.Pictures.Add(new Picture
            {
                sizeY = 0
            });
            da.Generate(dto);




            da.word.Model.path = @"c:\\aaa.worx";
            da.word.Models.Add(new pikunword_model {
                rId = 1,
                execut_type = pikun_execut_type.picture,
                picture = new Picture { sizeX = 0 },
            });
            da.word.Models.Add(new pikunword_model
            {
                rId = 2,
                execut_type = pikun_execut_type.paragraphs,
                paragraph = new Paragraph { align = "center" },
            });
        }
    }
}
