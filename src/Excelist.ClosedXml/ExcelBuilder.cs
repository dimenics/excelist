using System.Reflection;
using ClosedXML.Excel;
using ClosedXML.Graphics;

namespace System.Collections.Generic
{
    internal class ExcelBuilder<T>
    {
        private readonly IEnumerable<T> _collection;
        private readonly ExcelSettings _settings;

        private readonly XLWorkbook _workBook;
        private readonly IXLWorksheet _worksheet;

        internal ExcelBuilder(IEnumerable<T> collection, ExcelSettings settings)
        {
            LoadOptions.DefaultGraphicEngine = new DefaultGraphicEngine(settings.Font);

            _collection = collection;
            _settings = settings;

            _workBook = new XLWorkbook();
            _worksheet = _workBook.Worksheets.Add(_settings.SheetName);
        }

        internal ExcelBuilder<T> CreateHeaders()
        {
            PropertyInfo[] properties = typeof(T).GetProperties();

            char colNo = 'A';
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo? property = properties[i];
                string col = colNo + "1";
                _worksheet.Cell(col).Value = property?.Name ?? "";

                if (i < properties.Length - 1)
                    colNo++;
            }

            IXLRange rngTable = _worksheet.Range($"A1:{colNo}1");
            rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngTable.Style.Font.Bold = true;
            rngTable.Style.Font.FontColor = _settings.Color;
            rngTable.Style.Fill.BackgroundColor = _settings.BackgroundColor;

            return this;
        }

        internal ExcelBuilder<T> CreateRows()
        {
            _worksheet.Cell("A2").InsertData(_collection);

            return this;
        }

        internal ExcelBuilder<T> Conclude()
        {
            _worksheet.Columns().AdjustToContents();
            return this;
        }

        internal MemoryStream Build()
        {
            var ms = new MemoryStream();
            _workBook.SaveAs(ms);
            _workBook.Dispose();

            ms.Position = 0;
            return ms;
        }
    }
}