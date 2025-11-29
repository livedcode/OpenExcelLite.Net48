using Xunit;
using OpenExcelLite.Net48;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenExcelLite.Net48.Tests
{
    public class SheetTests
    {
        [Fact]
        public void AddRow_ShouldIncreaseRowCount()
        {
            var wb = new ExcelWorkbook();
            var sheet = wb.AddSheet("Data");

            sheet.AddRow("A", "B", "C");
            sheet.AddRow("1", "2", "3");

            Assert.Equal(2, sheet.RowCount);
            Assert.Equal(3, sheet.ColumnCount);
        }

        [Fact]
        public void AutoColumnWidth_ShouldApplyWidths()
        {
            var wb = new ExcelWorkbook();
            var sheet = wb.AddSheet("Data");

            sheet.AddRow("Short", "This is a long string for width test");

            var bytes = wb.ToArray();

            using (var ms = new MemoryStream(bytes))
            {
                using (var doc = SpreadsheetDocument.Open(ms, false))
                {
                    var ws = doc.WorkbookPart.WorksheetParts.First().Worksheet;
                    var cols = ws.GetFirstChild<Columns>();

                    Assert.NotNull(cols);
                    Assert.True(cols.ChildElements.Count > 0);
                }
            }
        }
    }
}
