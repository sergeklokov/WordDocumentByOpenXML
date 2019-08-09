using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
//using DocumentFormat.OpenXml.Spreadsheet;  // we don't work with Excel in this example

namespace WordDocumentByOpenXML
{
    /// <summary>
    /// 1. Creating Word DOCX
    /// 2. Creating Word DOCX based on template
    /// 3. Export text data to the Word DOCX document 
    /// 4. Export images the Word DOCX document 
    /// 
    /// Serge Klokov 2019
    /// 
    /// URL below are used:
    /// 
    /// 
    /// </summary>
    class Program
    {
        const string templateFileName = "Template.dotx";
        const string docFileNameWTemplate = "ExportedBasedOnTemplate.docx";
        const string docFileName = "ExportedData.docx";

        // key values on the left if present in the Word Template, 
        // will be replaced by text from values
        static Dictionary<string, string> textData = new Dictionary<string, string>() {
            ["TextDataName"] = "Sergiy Klokov",
            ["TextDataAddress"] = "6200 E Germann Rd, Chandler, AZ",
            ["TextDataToday"] = DateTime.Now.ToLongDateString(),
            ["TextDataDOB"] = "10/03/1800",
            //[""] = "",
        };

        // key values on the left, if present,
        // will be replaced with images
        // we may set image size (TODO: image resizing in the word..)
        static Dictionary<string, string> imageData = new Dictionary<string, string>()
        {
            ["ImageDataPhoto"] = "",
            ["ImageDataGraph"] = "",
            ["ImageDataID"] = "",
            ["ImageDataBigPicture"] = "",
            ["ImageDataSmallPicture"] = "",
            //[""] = "",
        };

        static void Main(string[] args)
        {
            CreateWordDocPlain(docFileName);
            CreateWordDocBasedOnTemplate(templateFileName, docFileNameWTemplate);
             Console.WriteLine("");
        }

        private static void CreateWordDocPlain(string docFileName)
        {
            using (var wordDocument = WordprocessingDocument.Create(docFileName, WordprocessingDocumentType.Document)) {
            }


        }

        private static void CreateWordDocBasedOnTemplate(string templateFileName, string docFileNameWTemplate)
        {
            throw new NotImplementedException();
        }
    }
}
