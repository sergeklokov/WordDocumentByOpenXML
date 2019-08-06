using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordDocumentByOpenXML
{
    /// <summary>
    /// TODO:
    /// Export to the Word DOCX document
    /// Based on tebmplate with tables
    /// 
    /// Serge Klokov 2019
    /// </summary>
    class Program
    {
        const string fileName = "test.docx";
        static void Main(string[] args)
        {
            using (var wordDocument = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            Console.WriteLine("");
        }
    }
}
