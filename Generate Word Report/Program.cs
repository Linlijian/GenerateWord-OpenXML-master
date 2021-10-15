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
        }
    }
}
