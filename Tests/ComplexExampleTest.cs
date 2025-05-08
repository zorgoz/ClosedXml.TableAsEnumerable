using ClosedXml.Extensions;
using ClosedXML.Excel;
using System.Reflection;

namespace TESTS;

/// <summary>
/// Summary description for ComplexExampleTest
/// </summary>
[TestClass]
public class ComplexExampleTest
{
    public static XLWorkbook workbook;

    /// <summary>
    /// Initializes EPPLus excelPackage with the embedded content
    /// </summary>
    [ClassInitialize()]
    public static void MyClassInitialize(TestContext testContext)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = assembly.GetManifestResourceNames().First();

        using (Stream stream = assembly.GetManifestResourceStream(resourceName))
        {
            workbook = new XLWorkbook(stream);
        }
    }

    /// <summary>
    /// Frees up excelPackage
    /// </summary>
    [ClassCleanup()]
    public static void MyClassCleanup()
    {
        workbook.Dispose();
    }

    [TestMethod]
    public void TestComplexFixtures()
    {
        Assert.IsNotNull(workbook, "Excel package is null");

        // TEST3

        var workSheet = workbook.Worksheet("TEST3");
        Assert.IsNotNull(workSheet, "Worksheet TEST3 missing");

        var table = workSheet.Table("TEST3");
        Assert.IsNotNull(table, "Table TEST3 missing");

        Assert.IsTrue(table.RangeAddress.ToStringRelative() == "B2:G8", "Table3 is not as expected");
    }

    enum Manufacturers { Opel = 1, Ford, Toyota };
    class Cars
    {
        [ExcelTableColumn(ColumnIndex = 1)]
        public string LicensePlate { get; set; }

        [ExcelTableColumn]
        public Manufacturers Manufacturer { get; set; }

        [ExcelTableColumn(ColumnName = "Manufacturing date")]
        public DateTime? ManufacturingDate { get; set; }

        [ExcelTableColumn]
        public int Price { get; set; }

        [ExcelTableColumn]
        public ConsoleColor Color { get; set; }

        [ExcelTableColumn(ColumnName = "Is ready for traffic?")]
        public bool Ready { get; set; }

        public string unmappedProperty { get; set; }
        public override string ToString()
        {
            return $"{(Color.ToString())} {(Manufacturer.ToString())} {(ManufacturingDate?.ToShortDateString())}";
        }
    }

    [TestMethod]
    public void TestComplexExample()
    {
        var table = workbook.Table("TEST3");

        IEnumerable<Cars> enumerable = table.AsEnumerable<Cars>();
        IList<Cars> list = null;

        Assert.IsNotNull(enumerable);
        list = enumerable.ToList();

        Assert.IsTrue(list.Count() == 5, "We have 5 rows");
        Assert.IsTrue(list.DistinctBy(x => x.Color).Count() == 5, "We have 5 different colors");
        Assert.IsTrue(list.Count(x => string.IsNullOrWhiteSpace(x.LicensePlate)) == 1, "There is one without license plate");
        Assert.IsTrue(list.All(x => x.Manufacturer > 0), "All should have manufacturers");
        Assert.IsNull(list.Last().ManufacturingDate, "The last one's manufacturing date is unknown");
        Assert.IsTrue(list.Count(x => x.ManufacturingDate == null) == 1, "Only one manufacturig date is unknown");
        Assert.AreSame(list.Single(x => x.LicensePlate == null), list.Single(x => !x.Ready), "The one without the license plate is not ready");
        Assert.IsTrue(list.Max(x => x.Price) == 12000, "Highest price is 12000");
        Assert.AreEqual(new DateTime(2015, 3, 10), list.Max(x => x.ManufacturingDate), "Oldest was manufactured on 2015.03.10");
    }

}

