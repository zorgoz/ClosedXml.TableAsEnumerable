using System.Reflection;
using ClosedXml.Extensions;
using ClosedXML.Excel;

namespace TESTS;

[TestClass]
public class TableValidateTests
{
    private TestContext testContextInstance;
    private static XLWorkbook workBook;

    /// <summary>
    ///Gets or sets the test context which provides
    ///information about and functionality for the current test run.
    ///</summary>
    public TestContext TestContext
    {
        get
        {
            return testContextInstance;
        }
        set
        {
            testContextInstance = value;
        }
    }

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
            workBook = new XLWorkbook(stream);
        }
    }

    /// <summary>
    /// Frees up excelPackage
    /// </summary>
    [ClassCleanup()]
    public static void MyClassCleanup()
    {
        workBook.Dispose();
    }

    enum Manufacturers { Opel = 1, Ford, Mercedes};
    class WrongCars
    {
        [ExcelTableColumn(ColumnName = "License plate")]
        public required string LicensePlate { get; set; }

        [ExcelTableColumn]
        public Manufacturers Manufacturer { get; set; }

        [ExcelTableColumn(ColumnName = "Manufacturing date")]
        public DateTime ManufacturingDate { get; set; }

        [ExcelTableColumn(ColumnName = "Is ready for traffic?")]
        public bool Ready { get; set; }
    }

    [TestMethod]
    public void Test_TableValidation()
    {
        var table = workBook.Table("TEST3");

        Assert.IsNotNull(table, "We have TEST3 table");

        var validation = table.Validate<WrongCars>().ToList();

        Assert.IsNotNull(validation, "we have errors here");
        Assert.AreEqual(3, validation.Count, "We have 3 errors");
        Assert.IsTrue(validation.Exists(x => x.cellAddress.ToStringRelative().Equals("C6", StringComparison.InvariantCultureIgnoreCase)), "Toyota is not in the enumeration");
        Assert.IsTrue(validation.Exists(x => x.cellAddress.ToStringRelative().Equals("D7", StringComparison.InvariantCultureIgnoreCase)), "Date is null");
        Assert.IsTrue(validation.Exists(x => x.cellAddress.ToStringRelative().Equals("B6", StringComparison.InvariantCultureIgnoreCase)), "Required license plate missing");
    }
}
