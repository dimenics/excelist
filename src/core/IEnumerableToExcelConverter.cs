using System.IO;

namespace System.Collections.Generic
{
    public interface IEnumerableToExcelConverter<in T>
    {
        MemoryStream ToExcel(IEnumerable<T> collection);
    }
}