using Xunit;
using OpenExcelLite.Net48;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;

namespace OpenExcelLite.Net48.Tests
{
    public class WorkbookTests
    {
        [Fact]
        public void MultiSheet_Workbook_ShouldCreateAllSheets()
        {
            var wb = new ExcelWorkbook();
            wb.AddSheet("Users").AddRow("Id", "Name");
            wb.AddSheet("Orders").AddRow("OrderId", "Amount");

            var bytes = wb.ToArray();

            using (var ms = new MemoryStream(bytes))
            {
                using (var doc = SpreadsheetDocument.Open(ms, false))
                {
                    var sheets = doc.WorkbookPart.Workbook.Sheets;
                    Assert.Equal(2, sheets.Count());
                }
            }
        }
    }
}
