using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;

namespace OpenExcelLite.Net48
{
    /// <summary>
    /// Represents an Excel workbook with one or more sheets.
    /// Public entry point – matches OpenExcelLite v1.3.0 style.
    /// </summary>
    public sealed class ExcelWorkbook
    {
        private readonly List<ExcelSheet> _sheets = new List<ExcelSheet>();

        public IReadOnlyList<ExcelSheet> Sheets => _sheets.AsReadOnly();

        /// <summary>
        /// Add a new worksheet.
        /// </summary>
        public ExcelSheet AddSheet(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Sheet name cannot be null or empty.", nameof(name));

            var sheet = new ExcelSheet(this, name);
            _sheets.Add(sheet);
            return sheet;
        }

        /// <summary>
        /// Save workbook to a file path.
        /// </summary>
        public void SaveToFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be null or whitespace.", nameof(path));

            using (var fs = File.Create(path))
                SaveToStream(fs);
        }

        /// <summary>
        /// Save workbook to a writable stream.
        /// </summary>
        public void SaveToStream(Stream output)
        {
            if (output == null)
                throw new ArgumentNullException(nameof(output));
            if (!output.CanWrite)
                throw new ArgumentException("Stream must be writable.", nameof(output));

            using (var doc = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook, true))
            {
                var wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new Workbook();

                // Styles (header + date formats, etc.)
                var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = StylesheetFactory.CreateDefaultStylesheet();
                stylesPart.Stylesheet.Save();

                var sharedStrings = new SharedStringCache();
                var sheetsElem = wbPart.Workbook.AppendChild(new Sheets());

                uint sheetId = 1;

                foreach (var sheet in _sheets)
                {
                    var wsPart = wbPart.AddNewPart<WorksheetPart>();
                    var hyperlinks = new List<HyperlinkInfo>();

                    // Build sheet data + column widths
                    double[] columnWidths;
                    var sheetData = sheet.BuildSheetData(sharedStrings, hyperlinks, out columnWidths);

                    var worksheet = WorksheetWriter.BuildWorksheet(
                        wsPart,
                        sheetData,
                        hyperlinks,
                        columnWidths,
                        sheet.FreezeRowSplit,
                        sheet.FreezeColumnSplit);

                    wsPart.Worksheet = worksheet;
                    wsPart.Worksheet.Save();

                    sheetsElem.Append(new Sheet
                    {
                        Id = wbPart.GetIdOfPart(wsPart),
                        SheetId = sheetId++,
                        Name = sheet.Name
                    });
                }

                sharedStrings.WriteTo(wbPart);
                wbPart.Workbook.Save();
            }
        }

        /// <summary>
        /// Returns workbook as a byte array.
        /// </summary>
        public byte[] ToArray()
        {
            using (var ms = new MemoryStream())
            {
                SaveToStream(ms);
                return ms.ToArray();
            }
        }
    }
}
