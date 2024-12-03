namespace Excelist.ClosedXml.Tests
{
    [TestClass]
    public class ClosedXmlExcelExporterTests
    {
        [TestMethod]
        public void ToExcel_DefaultSettings_ShouldExport()
        {
            IEnumerableToExcelExporter<Customer> exporter = new EnumerableToExcelExporter<Customer>();

            List<Customer> customers =
            [
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
            ];

            MemoryStream memoryStream = exporter.ToExcel(customers);
            Assert.IsTrue(memoryStream.Length == 6582);
        }
    }
}