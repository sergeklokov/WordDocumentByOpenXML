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
    /// https://stackoverflow.com/questions/57531762/how-to-convert-word-document-created-from-template-by-openxml-into-memorystream/57532216
    /// 
    /// 
    /// </summary>
    public class Program
    {
        const string templateFileName = "Letter.docx";
        //const string templateFileName = "Template.dotx";
        const string docFileNameWTemplate = "ExportedBasedOnTemplate.docx";
        const string docFileName = "ExportedData.docx";

        // key values on the left if present in the Word Template, 
        // will be replaced by text from values
        static Dictionary<string, string> textData = new Dictionary<string, string>() {
            ["DataNameFrom"] = "Sergiy Klokov",
            ["DataToday"] = DateTime.Now.ToLongDateString(),
            ["DataDOB"] = "10/03/1800",
            ["DataAddressFrom"] = "200 S Cloud Rd",
            ["DataCityZipFrom"] = "Chandler, AZ, 85249",
            ["DataNameTo"] = "McKayla Klokov",
            ["DataAddressTo"] = "6200 E Germann Rd",
            ["DataCityZipTo"] = "Chandler, AZ, 85286",
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
            //CreateWordDocBasedOnTemplate(templateFileName, docFileNameWTemplate);
            Console.WriteLine("");
        }

        public static void CreateWordDocPlain(string docFileName)
        {
            using (var document = WordprocessingDocument.Create(docFileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart main = document.AddMainDocumentPart();
                main.Document = new Document();
                Body body = main.Document.AppendChild(new Body());

                //add text
                Paragraph p = body.AppendChild(new Paragraph());
                Run r = p.AppendChild(new Run());
                r.AppendChild(new Text("This the demo text in our demo document."));
                r.AppendChild(new Break());

                document.Close();
            }
        }

        public static void CreateWordDocBasedOnTemplate(string templateFileName, string docFileNameWTemplate)
        {
            using (var document = WordprocessingDocument.Create(docFileName, WordprocessingDocumentType.Document))
            {
                var body = document.MainDocumentPart.Document.Body;
                var paragraphs = body.Elements<Paragraph>();
                var text = paragraphs.SelectMany(p => p.Elements<Run>())
                    .SelectMany(r => r.Elements<Text>());

            }
        }

        public static MemoryStream GetWordDocumentFromTemplateWithTempFile()
        {
            string tempFileName = Path.GetTempFileName();
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"Controllers\" + templateFileName;

            using (var document = WordprocessingDocument.CreateFromTemplate(templatePath))
            {
                var body = document.MainDocumentPart.Document.Body;

                //add some text 
                Paragraph paraHeader = body.AppendChild(new Paragraph());
                Run run = paraHeader.AppendChild(new Run());
                run.AppendChild(new Text("This is body text"));

                OpenXmlPackage savedDoc = document.SaveAs(tempFileName); // Save result document, not modifying the template
                savedDoc.Close();  // can't read if it's open
                document.Close();
            }

            var memoryStream = new MemoryStream(File.ReadAllBytes(tempFileName)); // this works but I want to avoid saving and reading file

            //memoryStream.Position = 0; // should I rewind it? 
            return memoryStream;
        }


        /// <summary>
        /// Trick is to open template for editing, 
        /// then change type to the document
        /// Open method have returned stream
        /// 
        /// answer on 
        /// https://stackoverflow.com/questions/57531762/how-to-convert-word-document-created-from-template-by-openxml-into-memorystream/57532216
        /// </summary>
        /// <returns></returns>
        public static MemoryStream GetWordDocumentStreamFromTemplate()
        {
            var templatePath = AppDomain.CurrentDomain.BaseDirectory + @"Controllers\" + templateFileName;
            var memoryStream = new MemoryStream();

            using (var fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                fileStream.CopyTo(memoryStream);

            using (var document = WordprocessingDocument.Open(memoryStream, true))
            {
                document.ChangeDocumentType(WordprocessingDocumentType.Document); // change from template to document

                var body = document.MainDocumentPart.Document.Body;

                //add some text 
                Paragraph paraHeader = body.AppendChild(new Paragraph());
                Run run = paraHeader.AppendChild(new Run());
                run.AppendChild(new Text("This is body text"));

                document.Close();
            }

            memoryStream.Position = 0; //let's rewind it
            return memoryStream;
        }
    }


}
