using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class VLookupTests : TestBase
    {
        [TestMethod]
        public void VlookupShouldHandleWholeColumn()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["D1"].Value = 1;
                sheet.Cells["D2"].Value = 2;
                sheet.Cells["D3"].Value = 2;
                sheet.Cells["D4"].Value = 3;
                sheet.Cells["D5"].Value = 3;
                sheet.Cells["D6"].Value = 4;
                sheet.Cells["D7"].Value = 4;
                sheet.Cells["D8"].Value = 5;
                sheet.Cells["D9"].Value = 5;

                sheet.Cells["E1"].Value = "a";
                sheet.Cells["E2"].Value = "b";
                sheet.Cells["E3"].Value = "c";
                sheet.Cells["E4"].Value = "d";
                sheet.Cells["E5"].Value = "e";
                sheet.Cells["E6"].Value = "f";
                sheet.Cells["E7"].Value = "g";
                sheet.Cells["E8"].Value = "h";
                sheet.Cells["E9"].Value = "i";

                sheet.Cells["C10"].Formula = "VLOOKUP(3,D:E,2,FALSE)";
                sheet.Calculate();
                Assert.AreEqual("d", sheet.Cells["C10"].Value);
            }
        }


        [TestMethod]
        public void VlookupApprox_ByDate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["C1"].Formula = "TODAY()";

                sheet.Cells["A1"].Formula = "C1";
                sheet.Cells["A2"].Formula = "C1+1";
                sheet.Cells["A3"].Formula = "C1+3";
                sheet.Cells["A4"].Formula = "C1+7";

                sheet.Cells["B1"].Value = "a";
                sheet.Cells["B2"].Value = "b";
                sheet.Cells["B3"].Value = "c";
                sheet.Cells["B4"].Value = "d";

                sheet.Cells["D1"].Formula = "VLOOKUP(C1,A1:B4,2,TRUE)";
                sheet.Calculate();
                Assert.AreEqual("a", sheet.Cells["D1"].Value);
            }
        }

        [DataTestMethod]
        [DataRow(1, "a")]
        [DataRow(5, "d")]
        public void VlookupApprox_Find(int find, string expected)
        {
            using (var package = OpenPackage("VlookupApprox_Finds.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;

                sheet.Cells["B1"].Value = "a";
                sheet.Cells["B2"].Value = "b";
                sheet.Cells["B3"].Value = "c";
                sheet.Cells["B4"].Value = "d";

                sheet.Cells["D1"].Formula = $"VLOOKUP({find},A1:B4,2,TRUE)";
                sheet.Calculate();

                Assert.AreEqual(expected, sheet.Cells["D1"].Value);
                //SaveAndCleanup(package);
            }
        }

        [TestMethod]

        public void VlookupExact_NotFound()
        {
            using (var package = OpenPackage("VlookupExact_NotFound.xlsx",true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;

                sheet.Cells["B1"].Value = "a";
                sheet.Cells["B2"].Value = "b";
                sheet.Cells["B3"].Value = "c";
                sheet.Cells["B4"].Value = "d";

                sheet.Cells["D1"].Formula = $"VLOOKUP(5,A1:B4,2,FALSE)";
                sheet.Calculate();

                Assert.AreEqual(ErrorValues.NAError, sheet.Cells["D1"].Value);
                //SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void VlookupApprox_OutOfRangePositive_ReturnsRefError()
        {
            using (var package = OpenPackage("VlookupApprox_OutOfRangePositive_ReturnsRefError.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 4;

                sheet.Cells["C1"].Value = "a";
                sheet.Cells["C2"].Value = "b";
                sheet.Cells["C3"].Value = "c";
                sheet.Cells["C4"].Value = "d";

                sheet.Cells["D1"].Value = "aa";
                sheet.Cells["D2"].Value = "bb";
                sheet.Cells["D3"].Value = "cc";
                sheet.Cells["D4"].Value = "dd";

                sheet.Cells["E1"].Formula = $"VLOOKUP(2,B1:C4,{3},TRUE)"; // positive offset is out of range
                sheet.Calculate();

                Assert.AreEqual(ErrorValues.RefError, sheet.Cells["E1"].Value);

                //SaveAndCleanup(package);
            }
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(-1)]
        public void VlookupApprox_OutOfRangeNonPositive_ReturnsValueError(int offset)
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Value = 10;
                sheet.Cells["A2"].Value = 20;
                sheet.Cells["A3"].Value = 30;
                sheet.Cells["A4"].Value = 40;

                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["B4"].Value = 4;

                sheet.Cells["C1"].Value = "a";
                sheet.Cells["C2"].Value = "b";
                sheet.Cells["C3"].Value = "c";
                sheet.Cells["C4"].Value = "d";

                sheet.Cells["E1"].Formula = $"VLOOKUP(2,B1:C4,{offset},TRUE)";
                sheet.Calculate();

                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["E1"].Value);
                //SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ExactStrings()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("NewWs");

                sheet.Cells["A1"].Value = "a";
                sheet.Cells["A2"].Value = "b";
                sheet.Cells["A3"].Value = "c";
                sheet.Cells["A4"].Value = "d";

                sheet.Cells["B1"].Value = "aa";
                sheet.Cells["B2"].Value = "bb";
                sheet.Cells["B3"].Value = "cc";
                sheet.Cells["B4"].Value = "dd";

                sheet.Cells["C1"].Formula = $"VLOOKUP(\"c\", A1:B4, 2, FALSE)";
                sheet.Cells["C2"].Formula = $"VLOOKUP(\"d\", A1:B4, 2, FALSE)";

                sheet.Calculate();

                Assert.AreEqual("cc", sheet.Cells["C1"].Value);
                Assert.AreEqual("dd", sheet.Cells["C2"].Value);
            }
        }

        [TestMethod]
        public void ApproxStrings()
        {
            using (var package = OpenPackage("VLOOKUP_approxStrings.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("NewWs");

                sheet.Cells["A1"].Value = "a";
                sheet.Cells["A2"].Value = "b";
                sheet.Cells["A3"].Value = "c";
                sheet.Cells["A4"].Value = "d";

                sheet.Cells["B1"].Value = "aa";
                sheet.Cells["B2"].Value = "bb";
                sheet.Cells["B3"].Value = "cc";
                sheet.Cells["B4"].Value = "dd";

                //"easy" to find
                sheet.Cells["C1"].Formula = $"VLOOKUP(\"ca\", A1:B4, 2, TRUE)";
                //Slightly harder to find
                sheet.Cells["C2"].Formula = $"VLOOKUP(\"da\", A1:B4, 2, TRUE)";

                sheet.Calculate();
                Assert.AreEqual("cc", sheet.Cells["C1"].Value);
                Assert.AreEqual("dd", sheet.Cells["C2"].Value);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void VlookupApprox_Test()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("ws");

                sheet.Cells["F1"].Value = 5;

                var range = sheet.Cells["C1:D4"];

                range.Value = 3;

                sheet.Cells["C1"].Value = 4;
                sheet.Cells["C2"].Value = 5;

                sheet.Cells["F3"].Formula = "VLOOKUP(F1,C1:D4,2,TRUE)";

                sheet.Cells["F3"].Calculate();
                Assert.AreEqual(3, sheet.Cells["F3"].Value);
            }
        }

        //Potentially support?
        //[TestMethod]
        //public void ApproximateMixedTypesByDateNumberFormat()
        //{
        //    using (var package = OpenPackage("VlookupApprox_MixedTypesByDateNumberFormat.xlsx",true))
        //    {
        //        // STAGING 
        //        var sheet = package.Workbook.Worksheets.Add("test");
        //        // mimicking error scenario with date value to reference
        //        sheet.Cells["A1"].Formula = "TODAY()";

        //        // VLOOKUP INPUT
        //        sheet.Cells["F1"].Formula = "A1+1";
        //        sheet.Cells["F1"].Style.Numberformat.Format = "[$-409]mmmm\\ d\\,\\ yyyy;@";

        //        // RANGE
        //        // mimicking error scenario with very specific, mixed values and formats 
        //        sheet.Cells["C1"].Value = "Today"; // Vlookup returns #N/A with this literal string value in the range
        //        //sheet.Cells["C1"].Formula= "A1"; // Vlookup returns expected result with this Date value in the range
        //        sheet.Cells["C1"].Style.Numberformat.Format = "[$-409]mmm\\-yy;@";
        //        sheet.Cells["C2"].Formula = "A1+1";
        //        sheet.Cells["C2"].Style.Numberformat.Format = "mm-dd-yy";
        //        sheet.Cells["C3"].Formula = "A1+3";
        //        sheet.Cells["C3"].Style.Numberformat.Format = "mm-dd-yy";
        //        sheet.Cells["C4"].Formula = "A1+7";
        //        sheet.Cells["C4"].Style.Numberformat.Format = "mm-dd-yy";

        //        sheet.Cells["D1"].Value = ".01";
        //        sheet.Cells["D1"].Style.Numberformat.Format = "0%";
        //        sheet.Cells["D2"].Value = ".02";
        //        sheet.Cells["D2"].Style.Numberformat.Format = "0%";
        //        sheet.Cells["D3"].Value = ".03";
        //        sheet.Cells["D3"].Style.Numberformat.Format = "0%";
        //        sheet.Cells["D4"].Value = ".04";
        //        sheet.Cells["D4"].Style.Numberformat.Format = "0%";

        //        // VLOOKUP OUTPUT
        //        sheet.Cells["F3"].Formula = "VLOOKUP(F1,C1:D4,2,TRUE)";

        //        //var logfile = new FileInfo(@"c:\temp\logfile.txt");
        //        //package.Workbook.FormulaParserManager.AttachLogger(logfile);

        //        sheet.Calculate();

        //        //var range = sheet.Cells["C1:D4"];
        //        //var val = sheet.Cells["F1"].Value;

        //        Assert.AreEqual(".02", sheet.Cells["F3"].Value);

        //        SaveAndCleanup(package);

        //    }
        //}
    }
}
