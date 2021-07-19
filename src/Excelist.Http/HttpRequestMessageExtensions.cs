using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;

namespace System.Net.Http
{
    public static class HttpRequestMessageExtensions
    {
        public static HttpResponseMessage ExportToExcel<T>(this HttpRequestMessage request, IEnumerable<T> records, IEnumerableToExcelExporter<T> converter)
        {
            using MemoryStream stream = converter.ToExcel(records);

            ByteArrayContent content = new(stream.ToArray());
            HttpResponseMessage result = request.CreateResponse(HttpStatusCode.OK);
            result.Content = content;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = $"Export-{DateTime.Now:yyyy-MM-dd h:mm tt}.xlsx";

            return result;
        }
    }
}