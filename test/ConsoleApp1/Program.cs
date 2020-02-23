using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlSDKUtil.WordProcessing;
using System;
using System.Linq;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            ProcessWord();
            Console.ReadKey();
        }

        private static void ProcessWord()
        {
            var file = @"E:\GitHub\OpenXmlSDKUtil\test\ConsoleApp1\test.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(file, true))
            {
                var body = doc.MainDocumentPart.Document.Body;
                if(doc.HasStyleId("2"))
                {
                    Console.WriteLine("has styleId 2");
                }
                if (doc.HasStyleId("3"))
                {
                    Console.WriteLine("has styleId 3");
                }

                var styleNames = doc.MainDocumentPart.StyleDefinitionsPart.Styles.Descendants<StyleName>();
                if (styleNames.Any())
                {

                }
                var table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                TableRow row = table.Elements<TableRow>().ElementAt(0);
                TableCell cell = row.Elements<TableCell>().ElementAt(0);
                var pic = cell.Descendants<PIC.Picture>().FirstOrDefault();
                if(pic!=null)
                {
                    var imageFile = @"E:\GitHub\OpenXmlSDKUtil\test\ConsoleApp1\tmp.jpg";
                    var relationshipId = doc.AddImage(imageFile);
                    pic.BlipFill.Blip.Embed.Value = relationshipId;
                }
                
            }
        }
    }
}
