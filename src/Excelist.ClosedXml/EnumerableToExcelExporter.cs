using System.IO;

namespace System.Collections.Generic
{
    public class EnumerableToExcelExporter<T> : IEnumerableToExcelExporter<T>
    {
        private readonly ExcelSettings _settings;

        public EnumerableToExcelExporter()
        {
            _settings = new ExcelSettings();
        }

        public EnumerableToExcelExporter(ExcelSettings settings)
        {
            _settings = settings;
        }

        public MemoryStream ToExcel(IEnumerable<T> collection)
        {
            ExcelBuilder<T> builder = new(collection, _settings);
            var package = builder.CreateHeaders().CreateRows().Conclude().Build();
            return package;
        }
    }
}