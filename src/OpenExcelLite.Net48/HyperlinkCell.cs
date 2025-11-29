using System;

namespace OpenExcelLite.Net48
{
    /// <summary>
    /// Represents a hyperlink cell: display text + URL.
    /// </summary>
    public sealed class HyperlinkCell
    {
        public string Text { get; }
        public string Url { get; }

        public HyperlinkCell(string text, string url)
        {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            Url = url ?? throw new ArgumentNullException(nameof(url));
        }
    }
}
