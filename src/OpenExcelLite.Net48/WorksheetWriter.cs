using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelLite.Net48
{
    internal sealed class HyperlinkInfo
    {
        public string CellRef { get; }
        public string Url { get; }

        public HyperlinkInfo(string cellRef, string url)
        {
            CellRef = cellRef;
            Url = url;
        }
    }

    internal static class WorksheetWriter
    {
        /// <summary>
        /// Create a single cell, including style and value.
        /// </summary>
        public static Cell CreateCell(
            string cellRef,
            ExcelSheet.CellData data,
            SharedStringCache sharedStrings,
            IList<HyperlinkInfo> links,
            bool isHeaderRow)
        {
            var cell = new Cell { CellReference = cellRef };

            uint? styleIndex = null; // 0 = default, 1 = header, 2 = date

            switch (data.Type)
            {
                case ExcelSheet.ExcelCellType.Empty:
                    // leave value null
                    break;

                case ExcelSheet.ExcelCellType.Text:
                    {
                        string s = Convert.ToString(data.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                        int index = sharedStrings.GetOrAdd(s);
                        cell.DataType = CellValues.SharedString;
                        cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));
                        break;
                    }

                case ExcelSheet.ExcelCellType.Number:
                    {
                        string n = Convert.ToString(data.Value, CultureInfo.InvariantCulture) ?? "0";
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(n);
                        break;
                    }

                case ExcelSheet.ExcelCellType.Boolean:
                    {
                        bool b = Convert.ToBoolean(data.Value, CultureInfo.InvariantCulture);
                        cell.DataType = CellValues.Boolean;
                        cell.CellValue = new CellValue(b ? "1" : "0");
                        break;
                    }

                case ExcelSheet.ExcelCellType.DateTime:
                    {
                        double oa;
                        if (data.Value is DateTime dt)
                            oa = dt.ToOADate();
                        else if (data.Value is DateTimeOffset dto)
                            oa = dto.DateTime.ToOADate();
                        else
                        {
                            string s = Convert.ToString(data.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                            int idx = sharedStrings.GetOrAdd(s);
                            cell.DataType = CellValues.SharedString;
                            cell.CellValue = new CellValue(idx.ToString(CultureInfo.InvariantCulture));
                            return cell;
                        }

                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(oa.ToString(CultureInfo.InvariantCulture));
                        styleIndex = 2; // date style
                        break;
                    }

                case ExcelSheet.ExcelCellType.Hyperlink:
                    {
                        var h = (HyperlinkCell)data.Value;
                        int index = sharedStrings.GetOrAdd(h.Text ?? string.Empty);
                        cell.DataType = CellValues.SharedString;
                        cell.CellValue = new CellValue(index.ToString(CultureInfo.InvariantCulture));

                        links.Add(new HyperlinkInfo(cellRef, h.Url));
                        break;
                    }
            }

            if (isHeaderRow)
            {
                // header style overrides (bold + fill)
                styleIndex = 1;
            }

            if (styleIndex.HasValue)
                cell.StyleIndex = styleIndex.Value;

            return cell;
        }

        /// <summary>
        /// Build worksheet (columns, sheet views, data, hyperlinks).
        /// </summary>
        public static Worksheet BuildWorksheet(
            WorksheetPart wsPart,
            SheetData sheetData,
            IList<HyperlinkInfo> links,
            double[] columnWidths,
            uint? freezeRowSplit,
            uint? freezeColumnSplit)
        {
            var ws = new Worksheet();

            // Columns (for auto-fit)
            if (columnWidths != null && columnWidths.Length > 0)
            {
                var cols = new Columns();
                for (int i = 0; i < columnWidths.Length; i++)
                {
                    var col = new Column
                    {
                        Min = (uint)(i + 1),
                        Max = (uint)(i + 1),
                        Width = columnWidths[i],
                        CustomWidth = true
                    };
                    cols.Append(col);
                }

                ws.Append(cols);
            }

            // Freeze panes (SheetViews)
            if (freezeRowSplit.HasValue || freezeColumnSplit.HasValue)
            {
                var sheetViews = new SheetViews();
                var sheetView = new SheetView { WorkbookViewId = 0U };

                var pane = new Pane
                {
                    State = PaneStateValues.Frozen
                };

                if (freezeColumnSplit.HasValue && freezeColumnSplit.Value > 0)
                    pane.HorizontalSplit = freezeColumnSplit.Value;

                if (freezeRowSplit.HasValue && freezeRowSplit.Value > 0)
                    pane.VerticalSplit = freezeRowSplit.Value;

                // Top-left cell after freeze
                uint topRow = (freezeRowSplit ?? 0) + 1;
                int leftCol = (int)(freezeColumnSplit ?? 0);
                pane.TopLeftCell = CellReferenceHelper.GetCellReference(topRow, leftCol);

                pane.ActivePane = PaneValues.BottomLeft;

                sheetView.Append(pane);
                sheetViews.Append(sheetView);
                ws.Append(sheetViews);
            }

            ws.Append(sheetData);

            // Hyperlinks
            if (links.Count > 0)
            {
                var hlCollection = new Hyperlinks();

                foreach (var link in links)
                {
                    var rel = wsPart.AddHyperlinkRelationship(
                        new Uri(link.Url, UriKind.Absolute),
                        true);

                    hlCollection.Append(new Hyperlink
                    {
                        Reference = link.CellRef,
                        Id = rel.Id
                    });
                }

                ws.Append(hlCollection);
            }

            return ws;
        }
    }

    /// <summary>
    /// Utility to build Excel cell references (A1, B2, AA10, etc.).
    /// </summary>
    internal static class CellReferenceHelper
    {
        public static string GetCellReference(uint rowIndex, int columnIndexZeroBased)
        {
            var columnName = GetColumnName(columnIndexZeroBased);
            return columnName + rowIndex.ToString(CultureInfo.InvariantCulture);
        }

        public static string GetColumnName(int columnIndexZeroBased)
        {
            int dividend = columnIndexZeroBased + 1;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                char letter = (char)('A' + modulo);
                columnName = letter + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }

    /// <summary>
    /// Creates a simple Stylesheet with:
    ///  - index 0: default
    ///  - index 1: header (bold + grey fill)
    ///  - index 2: date (NumberFormatId 14)
    /// </summary>
    internal static class StylesheetFactory
    {
        public static Stylesheet CreateDefaultStylesheet()
        {
            var fonts = new Fonts(
                new Font(),                          // 0: default
                new Font(new Bold())                 // 1: header bold
            );

            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),      // 0
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),   // 1 (required)
                new Fill(
                    new PatternFill(
                        new ForegroundColor { Rgb = HexBinaryValue.FromString("FFD9D9D9") })
                    { PatternType = PatternValues.Solid })                           // 2: header fill
            );

            var borders = new Borders(
                new Border()  // default border
            );

            var cellStyleFormats = new CellStyleFormats(
                new CellFormat()
            );

            var cellFormats = new CellFormats(
                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0
                }, // 0: default

                new CellFormat
                {
                    FontId = 1,
                    FillId = 2,
                    BorderId = 0,
                    ApplyFont = true,
                    ApplyFill = true
                }, // 1: header

                new CellFormat
                {
                    FontId = 0,
                    FillId = 0,
                    BorderId = 0,
                    NumberFormatId = 14U, // built-in short date
                    ApplyNumberFormat = true
                } // 2: date
            );

            var stylesheet = new Stylesheet
            {
                Fonts = fonts,
                Fills = fills,
                Borders = borders,
                CellStyleFormats = cellStyleFormats,
                CellFormats = cellFormats
            };

            return stylesheet;
        }
    }
}
