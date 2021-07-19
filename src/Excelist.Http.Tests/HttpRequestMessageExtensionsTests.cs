using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Net.Http;

namespace Excelist.Http.Tests
{
    [TestClass]
    public class HttpRequestMessageExtensionsTests
    {
        [TestMethod]
        public void ExportToExcel()
        {
            List<Customer> customers = new List<Customer>()
            {
                new Customer()
                {
                    Name = "Beverley Wyndhame",
                    Address = "1830 Metz Lane, Pennsville, New Jersey, USA",
                    Email = "bwyndhame@contoso.com"
                },

                new Customer()
                {
                    Name = "Brighton Buckley",
                    Address = "Lodorp 3113, Maffe, Namur, Belgium",
                    Email = "bbuckley@contoso.com"
                }
            };

            IEnumerableToExcelExporter<Customer> exporter = new OpenOfficeEnumerableToExcelExporter<Customer>();

            HttpRequestMessage request = new HttpRequestMessage();
            HttpResponseMessage response = request.ExportToExcel(customers, exporter);

            Assert.IsTrue(response.Content.Headers.ContentLength == 2982);
        }
    }
}
