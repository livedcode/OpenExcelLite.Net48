using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelLite.Net48
{
    /// <summary>
    /// Represents a single worksheet within a workbook.
    /// </summary>
    public sealed class ExcelSheet
    {
        private readonly ExcelWorkbook _workbook;
        private readonly List<RowData> _rows = new List<RowData>();

        internal string Name { get; }

        public int ColumnCount { get; private set; }
        public int RowCount => _rows.Count;

        internal bool AutoFitColumnsEnabled { get; private set; } = true;

        internal uint? FreezeRowSplit { get; private set; }
        internal uint? FreezeColumnSplit { get; private set; }

        internal ExcelSheet(ExcelWorkbook workbook, string name)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        // ----------------------------------------------------------------
        // PUBLIC API – Feature-parity with OpenExcelLite v1.3.0
        // ----------------------------------------------------------------

        /// <summary>
        /// Add a row from params.
        /// </summary>
        public ExcelSheet AddRow(params object[] values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            EnsureSchema(values.Length);

            var cells = new List<CellData>(values.Length);
            foreach (var v in values)
                cells.Add(CellData.FromObject(v));

            _rows.Add(new RowData(cells));
            return this;
        }

        /// <summary>
        /// Add a row from any enumerable (List&lt;object&gt;, etc.).
        /// </summary>
        public ExcelSheet AddRow(IEnumerable<object> values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));

            if (values is object[] arr)
                return AddRow(arr);

            var list = new List<object>();
            foreach (var v in values)
                list.Add(v);

            return AddRow(list.ToArray());
        }

        /// <summary>
        /// Add multiple rows.
        /// </summary>
        public ExcelSheet AddRows(IEnumerable<object[]> rows)
        {
            if (rows == null)
                throw new ArgumentNullException(nameof(rows));

            foreach (var row in rows)
            {
                if (row == null)
                    throw new ArgumentException("Row cannot be null in AddRows.", nameof(rows));

                AddRow(row);
            }

            return this;
        }

        /// <summary>
        /// Add N empty rows (all blank cells).
        /// </summary>
        public ExcelSheet AddEmptyRows(int count)
        {
            if (count < 0)
                throw new ArgumentOutOfRangeException(nameof(count));

            if (count == 0)
                return this;

            // If no schema exists yet, allow empty rows added BEFORE header.
            // They will have 0 columns temporarily.
            if (ColumnCount == 0)
            {
                for (int i = 0; i < count; i++)
                {
                    _rows.Add(new RowData(new List<CellData>())); // 0-column empty row
                }
                return this;
            }

            // Schema exists → fill correct number of empty cells
            for (int i = 0; i < count; i++)
            {
                var cells = new List<CellData>(ColumnCount);
                for (int c = 0; c < ColumnCount; c++)
                    cells.Add(CellData.Empty());

                _rows.Add(new RowData(cells));
            }

            return this;
        }

        /// <summary>
        /// Enable or disable simple auto column width (based on content length).
        /// Enabled by default.
        /// </summary>
        public ExcelSheet AutoFitColumns(bool enabled = true)
        {
            AutoFitColumnsEnabled = enabled;
            return this;
        }

        /// <summary>
        /// Freeze panes: split rows/columns. Example: FreezePanes(1) to freeze top row.
        /// </summary>
        public ExcelSheet FreezePanes(uint freezeRowCount, uint freezeColumnCount = 0)
        {
            FreezeRowSplit = freezeRowCount;
            FreezeColumnSplit = freezeColumnCount;
            return this;
        }

        // ----------------------------------------------------------------
        // INTERNAL – Called by ExcelWorkbook
        // ----------------------------------------------------------------

        internal SheetData BuildSheetData(SharedStringCache cache, List<HyperlinkInfo> links, out double[] columnWidths)
        {
            if (cache == null) throw new ArgumentNullException(nameof(cache));
            if (links == null) throw new ArgumentNullException(nameof(links));

            var sheetData = new SheetData();

            double[] widths = null;
            if (AutoFitColumnsEnabled && ColumnCount > 0)
                widths = new double[ColumnCount];

            uint rowIndex = 1;

            foreach (var rowData in _rows)
            {
                var row = new Row { RowIndex = rowIndex };
                bool isHeaderRow = (rowIndex == 1);

                for (int col = 0; col < rowData.Cells.Count; col++)
                {
                    string cellRef = CellReferenceHelper.GetCellReference(rowIndex, col);
                    var cellData = rowData.Cells[col];

                    // track width
                    if (widths != null)
                    {
                        string display = cellData.GetDisplayText();
                        int len = string.IsNullOrEmpty(display) ? 0 : display.Length;
                        if (len > 0)
                        {
                            double approxWidth = len + 2; // simple padding
                            if (approxWidth > widths[col])
                                widths[col] = approxWidth;
                        }
                    }

                    var cell = WorksheetWriter.CreateCell(
                        cellRef,
                        cellData,
                        cache,
                        links,
                        isHeaderRow);

                    row.Append(cell);
                }

                sheetData.Append(row);
                rowIndex++;
            }

            // Avoid zero width columns
            if (widths != null)
            {
                for (int i = 0; i < widths.Length; i++)
                {
                    if (widths[i] <= 0)
                        widths[i] = 8.0;
                }
            }

            columnWidths = widths;
            return sheetData;
        }

        // ----------------------------------------------------------------
        // Internal helpers + data structures
        // ----------------------------------------------------------------

        private void EnsureSchema(int columnCount)
        {
            if (columnCount <= 0)
                throw new ArgumentException("Row must have at least one column.", nameof(columnCount));

            if (ColumnCount == 0)
            {
                ColumnCount = columnCount;
            }
            else if (ColumnCount != columnCount)
            {
                throw new InvalidOperationException(
                    $"Row has {columnCount} cells, but the sheet schema requires {ColumnCount}.");
            }
        }

        internal sealed class RowData
        {
            public List<CellData> Cells { get; }

            public RowData(List<CellData> cells)
            {
                Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            }
        }

        internal enum ExcelCellType
        {
            Empty,
            Text,
            Number,
            Boolean,
            DateTime,
            Hyperlink
        }

        internal sealed class CellData
        {
            public ExcelCellType Type { get; }
            public object Value { get; }

            public CellData(ExcelCellType type, object value)
            {
                Type = type;
                Value = value;
            }

            public static CellData Empty() => new CellData(ExcelCellType.Empty, null);

            public static CellData FromObject(object value)
            {
                if (value == null || value is DBNull)
                    return Empty();

                if (value is HyperlinkCell h)
                    return new CellData(ExcelCellType.Hyperlink, h);

                if (value is string)
                    return new CellData(ExcelCellType.Text, value);

                if (value is bool)
                    return new CellData(ExcelCellType.Boolean, value);

                if (value is DateTime || value is DateTimeOffset)
                    return new CellData(ExcelCellType.DateTime, value);

                if (IsNumeric(value))
                    return new CellData(ExcelCellType.Number, value);

                return new CellData(ExcelCellType.Text, value.ToString());
            }

            public string GetDisplayText()
            {
                switch (Type)
                {
                    case ExcelCellType.Empty:
                        return string.Empty;

                    case ExcelCellType.Text:
                        return Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty;

                    case ExcelCellType.Number:
                        return Convert.ToString(Value, CultureInfo.InvariantCulture) ?? string.Empty;

                    case ExcelCellType.Boolean:
                        return Convert.ToBoolean(Value, CultureInfo.InvariantCulture) ? "TRUE" : "FALSE";

                    case ExcelCellType.DateTime:
                        if (Value is DateTime dt)
                            return dt.ToShortDateString();
                        if (Value is DateTimeOffset dto)
                            return dto.Date.ToShortDateString();
                        return Value.ToString();

                    case ExcelCellType.Hyperlink:
                        var h = (HyperlinkCell)Value;
                        return h.Text ?? string.Empty;

                    default:
                        return Value?.ToString() ?? string.Empty;
                }
            }

            private static bool IsNumeric(object v) =>
                v is sbyte || v is byte ||
                v is short || v is ushort ||
                v is int || v is uint ||
                v is long || v is ulong ||
                v is float || v is double || v is decimal;
        }
    }
}
