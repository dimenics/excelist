using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace System.Collections.Generic
{
    internal class ExcelBuilder<T>
    {
        private readonly IEnumerable<T> _collection;
        private readonly ExcelSettings _settings;

        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _worksheet;

        internal ExcelBuilder(IEnumerable<T> collection, ExcelSettings settings)
        {
            _collection = collection;
            _settings = settings;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add(_settings.SheetName);
        }

        internal ExcelBuilder<T> CreateHeaders()
        {
            PropertyInfo[] properties = typeof(T).GetProperties();

            using (ExcelRange range = _worksheet.Cells[1, 1, 1, properties.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(_settings.BackgroundColor);
                range.Style.Font.Color.SetColor(_settings.Color);
            }

            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties.ElementAt(i);
                _worksheet.Cells[1, i + 1].Value = property.Name;
            }

            return this;
        }

        internal ExcelBuilder<T> CreateRows()
        {
            PropertyInfo[] properties = typeof(T).GetProperties();
            int rowNo = 2;
            for (int i = 0; i < _collection.Count(); i++)
            {
                T record = _collection.ElementAt(i);
                for (int y = 0; y < properties.Length; y++)
                {
                    _worksheet.Cells[rowNo, y + 1].Value = record.GetPropValue(properties.ElementAt(y).Name)?.ToString() ?? "";
                }

                rowNo += 1;
            }

            return this;
        }

        internal ExcelBuilder<T> Conclude()
        {
            _worksheet.Cells[_worksheet.Dimension.Address].AutoFitColumns();
            return this;
        }

        internal ExcelPackage Build() => _package;
    }
}