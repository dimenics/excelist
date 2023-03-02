using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace Excelist.OpenXml.Tests
{

    [TestClass]
    public class OpenOfficeEnumerableToExcelExporterTests
    {
        [TestMethod]
        public void ToExcel_DefaultSettings_ShouldExport()
        {
            IEnumerableToExcelExporter<Customer> exporter = new OpenOfficeEnumerableToExcelExporter<Customer>();

            List<Customer> customers = new()
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

            MemoryStream memoryStream = exporter.ToExcel(customers);
            Assert.IsTrue(memoryStream.Length == 2982);
        }

        [TestMethod]
        public void ToExcel_CustomSettings_ShouldExport()
        {
            IEnumerableToExcelExporter<Customer> exporter = new OpenOfficeEnumerableToExcelExporter<Customer>(
                new ExcelSettings()
                {
                    BackgroundColor = System.Drawing.Color.GreenYellow,
                    Color = System.Drawing.Color.Black,
                    SheetName = "Customers"
                });

            List<Customer> customers = new()
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

            MemoryStream memoryStream = exporter.ToExcel(customers);
            Assert.IsTrue(memoryStream.Length == 3000);
        }
    }
}
