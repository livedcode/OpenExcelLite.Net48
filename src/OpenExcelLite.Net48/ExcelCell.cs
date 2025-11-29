using System;

namespace OpenExcelLite.Net48
{
    /// <summary>
    /// Helper factory for strongly-typed cell values.
    /// This is mostly syntactic sugar for library consumers.
    /// </summary>
    public static class ExcelCell
    {
        public static object Text(string value) => value;
        public static object Number<T>(T value) where T : struct => value;
        public static object Bool(bool value) => value;
        public static object Date(DateTime date) => date;
        public static object Hyperlink(string text, string url) => new HyperlinkCell(text, url);
    }
}
