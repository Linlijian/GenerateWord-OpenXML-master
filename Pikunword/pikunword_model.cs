﻿using System;
using System.Collections.Generic;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace Pikunword
{
    public class pikunword_model
    {
        //public int rId { get; set; }
        public string execut_type { get; set; }
        public string path { get; set; }

        public ExtendedFileProperties extended_file_properties { get; set; }
        public PikunPicture picture { get; set; }
        public PikunParagraph paragraph { get; set; }
        public PikunTable table { get; set; }
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

    public class PikunParagraph
    {
        public int rId { get; set; }
        public string text { get; set; }
        public string txt_justification { get; set; } //ข้อความบางส่วน align
        public string font { get; set; }
        public int font_size { get; set; }
        public string justification { get; set; } //ข้อความทั้งหมด align ไปทางเดียวกัน
        public string color { get; set; }
        public string highlight { get; set; }
        public string prop { get; set; }
        public int numbering_level_reference { get; set; } //ย่อหน้า
        public int numbering_id { get; set; } // 1 numbering | 2 bullet
        public List<PikunParagraphManyProp> many_prop { get; set; }
    }

    public class PikunNumberingDefinitions
    {
        public string number_format_values { get; set; } //numberFormatValuesDecimalABC | numberFormatValuesDecimal
        public string[] numbering_type { get; set; } // "-", ".", "ü", "o", etc.. ไม่เกิน  9 ตัว
        public string font { get; set; }
    }

    public class PikunParagraphManyProp
    {
        public int rId { get; set; }
        public string text { get; set; }
        public string font { get; set; }
        public int font_size { get; set; }
        public string justification { get; set; }
        public string color { get; set; }
        public string highlight { get; set; }
        public string[] prop { get; set; }
    }

    public class PikunPicture
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

    public class PikunTable
    {
        public int rId { get; set; }
        public string table_style { get; set; }
        public string table_width { get; set; } //defualt = 0
        public bool have_table_cell_margin { get; set; } //tableCellMarginDefault
        public bool table_cell_width_auto { get; set; } //tableCellMarginDefault
        public List<PikunTableGrid> table_grid { get; set; }
    }

    public class PikunTableGrid
    {
        public int rId { get; set; }
        public string grid_column { get; set; }
        public string table_cell_width { get; set; }
        public List<PikunTableCellProperties> table_cell_properties { get; set; }
    }

    public class PikunTableCellProperties
    {
        public int rId { get; set; }
        public string text { get; set; }
        public string font { get; set; }
        public int font_size { get; set; }
        public string justification { get; set; }
        public string color { get; set; }
        public string highlight { get; set; }
        public string[] prop { get; set; }

        public string top_border_color { get; set; }
        public int top_border_size { get; set; }
        public int top_border_space { get; set; }
        public string left_border_color { get; set; }
        public int left_border_size { get; set; }
        public int left_border_space { get; set; }
        public string right_border_color { get; set; }
        public int right_border_size { get; set; }
        public int right_border_space { get; set; }
        public string bottom_border_color { get; set; }
        public int bottom_border_size { get; set; }
        public int bottom_border_space { get; set; }

        public string top_margin { get; set; }
        public string right_margin { get; set; }
        public string left_margin { get; set; }
        public string bottom_margin { get; set; }

        public string fill { get; set; }

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
