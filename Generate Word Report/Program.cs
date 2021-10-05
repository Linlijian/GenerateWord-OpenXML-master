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
            //gen2 G2 = new gen2();
            //G2.CreatePackage("gwn2.docx");

            //GeneratedClass G = new GeneratedClass();
            //G.CreatePackage("Test.docx");

            //Class1 G21 = new Class1();
            //G21.CreatePackage("G21.docx");

            Yona yona = new Yona();
            yona.CreatePackage("yona.docx");

            ColunmChart ColunmChart = new ColunmChart();
            ColunmChart.CreatePackage("ColunmChart.docx");

            Picture Picture = new Picture();
            Picture.CreatePackage("Picture.docx");

            Empty Empty = new Empty();
            Empty.CreatePackage("Empty_word.docx");

















            //Trying to generate PDF from Word
            /*
            object oMissing = System.Reflection.Missing.Value;
            // Use the dummy value as a placeholder for optional arguments
            mw.Document doc = word.Documents.Open("Test.docx", ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            object outputFileName = wordFile.FullName.Replace(".doc", ".pdf");
            object fileFormat = WdSaveFormat.wdFormatPDF;

            // Save document into PDF Format
            doc.SaveAs(ref outputFileName,
                ref fileFormat, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
            doc = null;*/

        }
    }
}
