using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelLite.Net48
{
    internal sealed class SharedStringCache
    {
        private readonly Dictionary<string, int> _lookup =
            new Dictionary<string, int>(StringComparer.Ordinal);

        private readonly List<string> _values = new List<string>();

        public int GetOrAdd(string text)
        {
            if (text == null) text = string.Empty;

            if (_lookup.TryGetValue(text, out int index))
                return index;

            index = _values.Count;
            _values.Add(text);
            _lookup[text] = index;
            return index;
        }

        public void WriteTo(WorkbookPart wbPart)
        {
            if (_values.Count == 0)
                return;

            var sstPart = wbPart.AddNewPart<SharedStringTablePart>();
            var table = new SharedStringTable();

            foreach (var v in _values)
                table.AppendChild(new SharedStringItem(new Text(v)));

            table.Count = (uint)_values.Count;
            table.UniqueCount = (uint)_values.Count;

            sstPart.SharedStringTable = table;
            sstPart.SharedStringTable.Save();
        }
    }
}
