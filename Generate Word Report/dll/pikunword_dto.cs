using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Generate_Word_Report.dll
{
    public class pikunword_dto : base_dto
    {
        public pikunword_dto()
        {
            Model = new pikunword_model();
            Models = new List<pikunword_model>();
            ExtendedFileProperties = new ExtendedFileProperties();
            Pictures = new List<PikunPicture>();
            Paragraphs = new List<PikunParagraph>();
            PackageProperties = new PackageProperties();
            NumberingDefinitions = new List<PikunNumberingDefinitions>();
        }
        public pikunword_model Model { get; set; }
        public List<pikunword_model> Models { get; set; }
        public ExtendedFileProperties ExtendedFileProperties { get; set; }
        public List<PikunPicture> Pictures { get; set; }
        public List<PikunParagraph> Paragraphs { get; set; }
        public PackageProperties PackageProperties { get; set; }
        public List<PikunNumberingDefinitions> NumberingDefinitions { get; set; }
    }

    public class pikun_execut_type
    {
        public const string create_packet = "create_packet";        
        public const string picture = "picture";
        public const string paragraphs = "paragraphs";
    }

    public class pikun_execut_function
    {
        public const string newLine = "newLine";
        public const string newLineNormal = "newLineNormal";
        public const string newLineManyprop = "newLineManyprop";
        public const string newLineNumbering = "newLineNumbering";
        public const string newLineImage = "newLineImage";
        public const string newLineImageNoFormat = "newLineImageNoFormat";
        public const string newLineNumberingProp = "newLineNumberingProp";
        public const string xxx = "xxx";
    }
}
