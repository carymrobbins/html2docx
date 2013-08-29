using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace html2docx
{
    class MainClass
    {
        public static void Main (string[] args)
        {
            if (args.Length != 2) {
                Console.WriteLine("Usage: mono {0} inputfile.html outputfile.docx",
				                  System.AppDomain.CurrentDomain.FriendlyName);
                return;
            }
            
            string inputFile = args[0];
            string outputFile = args[1];
            
            using (WordprocessingDocument myDoc =
			      	WordprocessingDocument.Create(
						outputFile, WordprocessingDocumentType.Document))
            {
                string altChunkId = "AltChunkId1";
                MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                mainPart.Document.Body = new Body();
                var chunk = mainPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.Html, altChunkId);
                using (FileStream fileStream =
                        File.Open(inputFile, FileMode.Open))
                {
                    chunk.FeedData(fileStream);
                }
                AltChunk altChunk = new AltChunk() {Id = altChunkId};
                mainPart.Document.Append(altChunk);
                mainPart.Document.Save();
            }
        }
    }
}
