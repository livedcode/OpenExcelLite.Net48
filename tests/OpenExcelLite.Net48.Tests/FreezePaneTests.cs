using Xunit;
using OpenExcelLite.Net48;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenExcelLite.Net48.Tests
{
    public class FreezePaneTests
    {
        [Fact]
        public void FreezeTopRow_ShouldCreatePane()
        {
            var wb = new ExcelWorkbook();
            var sheet = wb.AddSheet("Data");

            sheet.AddRow("Id", "Name");
            sheet.FreezePanes(1);

            var bytes = wb.ToArray();

            using (var ms = new MemoryStream(bytes))
            {
                using (var doc = SpreadsheetDocument.Open(ms, false))
                {
                    var ws = doc.WorkbookPart.WorksheetParts.First().Worksheet;
                    var views = ws.GetFirstChild<SheetViews>();

                    Assert.NotNull(views);
                    Assert.NotNull(views.Descendants<Pane>().FirstOrDefault());
                }
            }
        }
    }
}
