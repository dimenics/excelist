using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace System.Net.Http
{
    public static class HttpRequestMessageExtensions
    {
        private static object GetPropValue(this object src, string propName)
            => src.GetType().GetProperty(propName).GetValue(src, null);

        public static HttpResponseMessage ExportToExcel<T>(this HttpRequestMessage request, IEnumerable<T> records)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using ExcelPackage package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Logs");

            PropertyInfo[] properties = typeof(T).GetProperties();
            using (ExcelRange range = worksheet.Cells[1, 1, 1, properties.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.Purple);
                range.Style.Font.Color.SetColor(Color.White);
            }

            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties.ElementAt(i);
                worksheet.Cells[1, i + 1].Value = property.Name;
            }

            int rowNo = 2;
            for (int i = 0; i < records.Count(); i++)
            {
                T record = records.ElementAt(i);
                for (int y = 0; y < properties.Length; y++)
                {
                    worksheet.Cells[rowNo, y + 1].Value = record.GetPropValue(properties.ElementAt(y).Name)?.ToString() ?? "";
                }

                rowNo += 1;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            using MemoryStream stream = new MemoryStream(package.GetAsByteArray());
            ByteArrayContent content = new ByteArrayContent(stream.ToArray());
            HttpResponseMessage result = request.CreateResponse(HttpStatusCode.OK);
            result.Content = content;
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = $"Export-{DateTime.Now:yyyy-MM-dd h:mm tt}.xlsx";

            return result;
        }
    }
}