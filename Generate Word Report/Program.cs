using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using mw=Microsoft.Office.Interop.Word;

namespace Generate_Word_Report
{
    class Program
    {
        static void Main(string[] args)
        {
            Empty Empty = new Empty();
            Empty.CreatePackage("Empty_word.docx");
            

        }
    }
}
