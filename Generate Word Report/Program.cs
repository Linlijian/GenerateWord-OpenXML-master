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
            one_pic_format_size Empty = new one_pic_format_size();
            Empty.CreatePackage("one_pic_format_size.docx");
            

        }
    }
}
