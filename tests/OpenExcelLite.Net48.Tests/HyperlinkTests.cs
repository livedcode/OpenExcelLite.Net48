using Xunit;
using OpenExcelLite.Net48;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenExcelLite.Net48.Tests
{
    public class HyperlinkTests
    {
        [Fact]
        public void Hyperlink_ShouldCreateRelationship()
        {
            var wb = new ExcelWorkbook();
            var sheet = wb.AddSheet("Users");

            sheet.AddRow("Id", "Website");
            sheet.AddRow(1, new HyperlinkCell("OpenAI", "https://openai.com"));

            var bytes = wb.ToArray();

            using (var ms = new MemoryStream(bytes))
            {
                using (var doc = SpreadsheetDocument.Open(ms, false))
                {
                    var wsPart = doc.WorkbookPart.WorksheetParts.First();
                    var links = wsPart.Worksheet.Descendants<Hyperlink>().ToList();

                    Assert.Single(links);
                    Assert.Contains("openai.com", wsPart.HyperlinkRelationships.First().Uri.ToString());
                }
            }
        }
    }
}
