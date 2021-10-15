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
            Pictures = new List<Picture>();
            Paragraphs = new List<Paragraph>();
            PackageProperties = new PackageProperties();
        }

        public pikunword_model Model { get; set; }
        public List<pikunword_model> Models { get; set; }
        public ExtendedFileProperties ExtendedFileProperties { get; set; }
        public List<Picture> Pictures { get; set; }
        public List<Paragraph> Paragraphs { get; set; }
        public PackageProperties PackageProperties { get; set; }
    }

    public class pikun_execut_type
    {
        public const string create_packet = "create_packet";        
        public const string picture = "picture";
        public const string paragraphs = "paragraphs";
    }
}
