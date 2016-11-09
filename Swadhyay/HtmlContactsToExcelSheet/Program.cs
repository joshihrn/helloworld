using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlContactsToExcelSheet
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo d = new DirectoryInfo(@"C:\Users\v-hijos\Documents\ProjDocuments\SourceHtmlFiles");
            var dd = @"C:\Users\v-hijos\Documents\ProjDocuments\DestinationHtmlFiles";
            var excelFileName = @"C:\Users\v-hijos\Documents\ProjDocuments\DestinationHtmlFiles\YuvaContacts.xlsx";

            var excelFile = new FileInfo(excelFileName);


            // Create the file using the FileInfo object
            FileInfo[] files = d.GetFiles();
            foreach(FileInfo file in files)
            {

                using (StreamReader sr = file.OpenText())
                {
                    HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
                    document.LoadHtml(sr.ReadToEnd());
                    HtmlNode node = document.DocumentNode.SelectNodes("//div[@class='dataTables_scroll']")[0];
                    if(node != null)
                    {
                        string finalText = node.InnerHtml;

                        HtmlAgilityPack.HtmlDocument innerDocument = new HtmlAgilityPack.HtmlDocument();
                        innerDocument.LoadHtml(finalText);
                        var elementsWithStyleAttribute = innerDocument.DocumentNode.SelectNodes("//@style");

                        if (elementsWithStyleAttribute != null)
                        {
                            foreach (var element in elementsWithStyleAttribute)
                            {
                                HtmlAttribute attr = element.Attributes["style"];
                                var newStyle = CleanStyles(attr.Value);

                                element.Attributes.Remove(element.Attributes["style"]);
                                element.SetAttributeValue("style", newStyle);

                            }
                        }

                        File.WriteAllText(Path.Combine(dd, Path.GetFileNameWithoutExtension(file.Name) + ".html"), innerDocument.DocumentNode.InnerHtml);

                    }
                }
            }

            using (var package = new ExcelPackage(excelFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Yuva contacts");

                worksheet.Cells[1, 1].LoadFromText("First Name,Last Name,Address,City,State,ZIP,Age");

                DirectoryInfo destiNationDir = new DirectoryInfo(dd);

                FileInfo[] destFileInfos = destiNationDir.GetFiles("*.html");
                int rowCount = 2;
                foreach (FileInfo file in destFileInfos)
                {
                    using (StreamReader sr = file.OpenText())
                    {
                        HtmlAgilityPack.HtmlDocument document = new HtmlAgilityPack.HtmlDocument();
                        document.LoadHtml(sr.ReadToEnd());

                        int columnCount = 1;
                        foreach (HtmlNode col in document.DocumentNode.SelectNodes("//table//tr//td"))
                        {
                                worksheet.Cells[rowCount, columnCount].Value = col.InnerText;
                                columnCount++;
                                if(columnCount > 7)
                                {
                                    columnCount = 1;
                                    rowCount++;
                                }
                        }
                    }
                }
                package.Save();
            }

        }

        public static string CleanStyles(string oldStyles)
        {
            string newStyles = "";
            foreach (var entries in oldStyles.Split(';'))
            {
                var values = entries.Split(':');
                if(values[0].Trim().ToLower() != "height")
                {
                    newStyles += entries + ";";
                }
                else
                {
                    string foundnothing = "true;";
                }
            }
            return newStyles;
        }
    }
}
