
using ClosedXML.Excel;

namespace System.Collections.Generic
{
    public class ExcelSettings
    {
        public XLColor BackgroundColor { get; set; }

        public XLColor Color { get; set; }

        public string SheetName { get; set; } = "Sheet1";

        public string Font { get; set; } = "Monospace";
    }
}