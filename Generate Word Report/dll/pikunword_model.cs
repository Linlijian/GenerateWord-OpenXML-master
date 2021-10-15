using System;
using System.Collections.Generic;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace Generate_Word_Report.dll
{
    public class pikunword_model
    {
        public int rId { get; set; }
        public string execut_type { get; set; }
        public string path { get; set; }

        public ExtendedFileProperties extended_file_properties { get; set; }
        public Picture picture { get; set; }
        public Paragraph paragraph { get; set; }
        public PackageProperties package_properties { get; set; }
    }

    public class ExtendedFileProperties
    {
        public Ap.Template template { get; set; }
        public Ap.TotalTime totalTime { get; set; }
        public Ap.Pages pages { get; set; }
        public Ap.Words words { get; set; }
        public Ap.Characters characters { get; set; }
        public Ap.Application application { get; set; }
        public Ap.DocumentSecurity documentSecurity { get; set; }
        public Ap.Lines lines { get; set; }
        public Ap.Paragraphs paragraphs { get; set; }
        public Ap.ScaleCrop scaleCrop { get; set; }
        public Ap.Company company { get; set; }
        public Ap.LinksUpToDate linksUpToDate { get; set; }
        public Ap.CharactersWithSpaces charactersWithSpaces { get; set; }
        public Ap.SharedDocument sharedDocument { get; set; }
        public Ap.HyperlinksChanged hyperlinksChanged { get; set; }
        public Ap.ApplicationVersion applicationVersion { get; set; }
    }

    public class Paragraph
    {
        public int rId { get; set; }
        public string text { get; set; }
        public string align { get; set; }
        public string txt_align { get; set; }
        public string[] many_prop { get; set; }
    }

    public class Picture
    {
        public int rId { get; set; }
        public string base64image { get; set; }
        public int sizeX { get; set; }
        public int sizeY { get; set; }
        public string horizontal_alignment { get; set; }
        public string layout_option { get; set; } //format
        public int horizontal_position { get; set; }
        public int vertical_position { get; set; }
    }

    public class PackageProperties
    {
        public string creator { get; set; }
        public string title { get; set; }
        public string subject { get; set; }
        public string keywords { get; set; }
        public string description { get; set; }
        public string revision { get; set; }
        public DateTime created { get; set; }
        public DateTime modified { get; set; }
        public string lastModifiedBy { get; set; }
    }
    
}
