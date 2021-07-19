using System.IO;
using OfficeOpenXml;

namespace System.Collections.Generic
{
    public class OpenOfficeEnumerableToExcelExporter<T> : IEnumerableToExcelExporter<T>
    {
        private readonly ExcelSettings _settings;

        public OpenOfficeEnumerableToExcelExporter()
        {
            _settings = new ExcelSettings();
        }

        public OpenOfficeEnumerableToExcelExporter(ExcelSettings settings)
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