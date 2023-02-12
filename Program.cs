using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace ReplacePlaceholderWithJson
{
   class Program
   {
      static void Main(string[] args)
      {
         try
         {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open("existingDocument.docx", true))
            {
               MainDocumentPart mainPart = wordDoc.MainDocumentPart;
               string docText = mainPart.Document.Body.InnerText;

               JObject jsonData = JObject.Parse(File.ReadAllText("data.json"));

               Regex placeholderRegex = new Regex("{(.*?)}");
               MatchCollection placeholders = placeholderRegex.Matches(docText);

               foreach (Match placeholder in placeholders)
               {
                  string key = placeholder.Groups[1].Value;
                  JToken value = jsonData[key];
                  // Console.WriteLine("key", value);
                  if (value != null)
                  {
                     docText = docText.Replace("{" + key + "}", value.ToString());
                  }
               }

               mainPart.Document.Body.RemoveAllChildren();
               mainPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(docText))));

               wordDoc.Save();
               Console.WriteLine("Placeholders replaced successfully with data from the JSON file.");
            }
         }
         catch (Exception ex)
         {
            Console.WriteLine("An error occurred while replacing placeholders with data from the JSON file: " + ex.Message);
         }
      }
   }
}
