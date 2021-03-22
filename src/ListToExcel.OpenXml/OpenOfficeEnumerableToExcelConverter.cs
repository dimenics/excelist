using System.IO;
using OfficeOpenXml;

namespace System.Collections.Generic
{
    public class OpenOfficeEnumerableToExcelConverter<T> : IEnumerableToExcelConverter<T>
    {
        private readonly ExcelSettings _settings;

        public OpenOfficeEnumerableToExcelConverter()
        {
            _settings = new ExcelSettings();
        }

        public OpenOfficeEnumerableToExcelConverter(ExcelSettings settings)
        {
            _settings = settings;
        }

        public MemoryStream ToExcel(IEnumerable<T> collection)
        {
            ExcelBuilder<T> builder = new(collection, _settings);
            ExcelPackage package = builder.CreateHeaders().CreateRows().Conclude().Build();

            return new MemoryStream(package.GetAsByteArray());
        }
    }
}