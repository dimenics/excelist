using System.IO;
using OfficeOpenXml;

namespace System.Collections.Generic
{
    public class DefaultEnumerableToExcelConverter<T> : IEnumerableToExcelConverter<T>
    {
        public MemoryStream ToExcel(IEnumerable<T> collection)
        {
            ExcelBuilder<T> builder = new ExcelBuilder<T>(collection);
            ExcelPackage package = builder.CreateHeaders().CreateRows().Conclude().Build();

            return new MemoryStream(package.GetAsByteArray());
        }
    }
}