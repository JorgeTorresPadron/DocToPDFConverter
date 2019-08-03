﻿using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Office2010.ExcelAc;


namespace DocXToPdfConverter
{
    
    public class Placeholders
    {
        //NewLineTags are important only for .docx as input. If you use .html as input, then just use "<br

        public string NewLineTag { get; set; }



        //Start and End Tags can e. g. be both "##"
        //A placeholder could be ##TextPlaceHolder##
        public string TextPlaceholderStartTag { get; set; }
        public string TextPlaceholderEndTag { get; set; }

        public Dictionary<string,string> TextPlaceholders { get; set; }

        /*
         * For tables it works that way:
         * 1. If you have a table in the word document, create 1 row with a different Dictionary keys
         * Then e. g. you want to have 10 rows in the end, you add 10 values to each array of the Dictionary value
         *
         * A placeholder could be ==TextPlaceHolder==
         */
        //Start and End Tags can e. g. be both "=="

        public string TablePlaceholderStartTag { get; set; }
        public string TablePlaceholderEndTag { get; set; }
        public List<Dictionary<string, string[]>> TablePlaceholders { get; set; }


        /*
         * Important: The MemoryStream may carry an image.
         * Allowed file types: JPEG/JPG, BMP, TIFF, GIF, PNG
         */

        //Take different replacement tags here, else there may be collision with the text replacements,
        //e. g. "++" 
        public string ImagePlaceholderStartTag { get; set; }
        public string ImagePlaceholderEndTag { get; set; }

        public Dictionary<string, MemoryStream> ImagePlaceholders { get; set; }
    }
}