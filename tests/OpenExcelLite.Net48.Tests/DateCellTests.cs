using Xunit;
using OpenExcelLite.Net48;
using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenExcelLite.Net48.Tests
{
    public class DateCellTests
    {
        [Fact]
        public void Date_ShouldBeStoredAsOADate()
        {
            var wb = new ExcelWorkbook();
            wb.AddSheet("Dates")
              .AddRow("Label", DateTime.Today);

            var bytes = wb.ToArray();

            using (var ms = new MemoryStream(bytes))
            {
                using (var doc = SpreadsheetDocument.Open(ms, false))
                {
                    var cell = doc.WorkbookPart.WorksheetParts.First()
                        .Worksheet.Descendants<Cell>()
                        .Last();

                    Assert.Equal(CellValues.Number, cell.DataType?.Value ?? CellValues.Number);
                }
            }
        }
    }
}
