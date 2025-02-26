﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using OfficeOpenXml.Drawing;
using System.Linq;

namespace EPPlusTest.Drawing
{
    [TestClass]
    public class CopyDrawingTests : TestBase
    {
        //Sheet 1: 4, 0-3
        //Sheet 2: 9, 0-8
        //Sheet 4: 7, 0-6

        //Copy Shape Tests
        [TestMethod]
        public void CopyShapeSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            Assert.IsTrue(ws0.Drawings.Count < 5);
            ws0.Drawings[0].Copy(ws0, 25, 1);
            Assert.AreEqual(5, ws0._drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyShapeOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws0.Drawings[0].Copy(ws1, 10, 10);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws0.Drawings[0].Copy(ws, 10, 10);
            Assert.AreEqual(1, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }
        [TestMethod]
        public void CopyShapeBlipFillTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws0.Drawings[1].Copy(ws1, 10, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }

        //Copy Picture Tests
        [TestMethod]
        public void CopyPictureSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws1.Drawings[0].Copy(ws1, 0, 15);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyPictureOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws0.Drawings.Count < 5);
            ws1.Drawings[0].Copy(ws0, 20, 1);
            Assert.AreEqual(5, ws0.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyPictureOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws0 = p2.Workbook.Worksheets.Add("Sheet1");
            ws1.Drawings[0].Copy(ws0, 1, 1);
            Assert.AreEqual(1, ws0.Drawings.Count);
            SaveAndCleanup(p2);
        }

        //Copy Control Tests
        [TestMethod]
        public void CopyControlSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws1.Drawings[1].Copy(ws1, 25, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyControlOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws1.Drawings[1].Copy(ws2, 20, 1);
            Assert.AreEqual(8, ws2.Drawings.Count);
            ws1.Drawings[2].Copy(ws2, 40, 1);
            ws1.Drawings[1].Copy(ws2, 50, 1);
            Assert.AreEqual(10, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyControlOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws1 = p.Workbook.Worksheets[1];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws1.Drawings[1].Copy(ws, 20, 1);
            ws1.Drawings[2].Copy(ws, 40, 1);
            ws1.Drawings[1].Copy(ws, 50, 1);
            Assert.AreEqual(3, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }

        //Copy Slicer Tests
        [TestMethod]
        public void CopySlicerSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            Assert.IsTrue(ws0.Drawings.Count < 5);
            ws0.Drawings[2].Copy(ws0, 1, 25, 0, 0);
            Assert.AreEqual(5, ws0.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopySlicerOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            var ws2 = p.Workbook.Worksheets[2];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws0.Drawings[2].Copy(ws2, 1, 15, 0, 0);
            Assert.AreEqual(8, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopySlicerOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws0 = p.Workbook.Worksheets[0];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            var ex = Assert.ThrowsException<InvalidOperationException>(() => ws0.Drawings[2].Copy(ws, 1, 15, 0, 0));
            Assert.AreEqual("Table slicers can't be copied from one workbook to another.", ex.Message);
        }

        //Copy Chart Tests
        [TestMethod]
        public void CopyChartSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws2.Drawings[0].Copy(ws2, 20, 1);
            Assert.AreEqual(8, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyChartOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws2.Drawings[0].Copy(ws1, 20, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyChartOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws2.Drawings[0].Copy(ws, 20, 1);
            Assert.AreEqual(1, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }

        //Copy Group Shape Tests
        [TestMethod]
        public void CopyGroupShapeSameWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            Assert.IsTrue(ws2.Drawings.Count < 8);
            ws2.Drawings[1].Copy(ws2, 5, 20);
            Assert.AreEqual(8, ws2.Drawings.Count);
            ws2.Drawings[2].Copy(ws2, 5, 25);
            ws2.Drawings[4].Copy(ws2, 5, 30);
            ws2.Drawings[5].Copy(ws2, 5, 35);
            ws2.Drawings[6].Copy(ws2, 5, 40);
            Assert.AreEqual(12, ws2.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyGroupShapeOtherWorksheetTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            var ws1 = p.Workbook.Worksheets[1];
            Assert.IsTrue(ws1.Drawings.Count < 10);
            ws2.Drawings[1].Copy(ws1, 5, 20);
            Assert.AreEqual(10, ws1.Drawings.Count);
            ws2.Drawings[2].Copy(ws1, 5, 25);
            ws2.Drawings[4].Copy(ws1, 5, 30);
            ws2.Drawings[5].Copy(ws1, 5, 35);
            ws2.Drawings[6].Copy(ws1, 5, 40);
            Assert.AreEqual(14, ws1.Drawings.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void CopyGroupShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            ws2.Drawings[1].Copy(ws, 1, 1);
            ws2.Drawings[2].Copy(ws, 1, 5);
            ws2.Drawings[4].Copy(ws, 5, 10);
            ws2.Drawings[5].Copy(ws, 5, 15);
            Assert.AreEqual(4, ws.Drawings.Count);
            SaveAndCleanup(p2);
        }
        [TestMethod]
        public void CopySlicerInGroupShapeOtherWorkbookTest()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws2 = p.Workbook.Worksheets[2];
            using var p2 = OpenPackage("Target.xlsx", true);
            var ws = p2.Workbook.Worksheets.Add("Sheet1");
            var ex = Assert.ThrowsException<InvalidOperationException>(() => ws2.Drawings[6].Copy(ws, 5, 40));
            Assert.AreEqual("Table slicers can't be copied from one workbook to another.", ex.Message);
        }

        //Copy PivotTable Slicers
        [TestMethod]
        public void CopyPivotSlicerSameWorksheet()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws3 = p.Workbook.Worksheets[3];
            Assert.IsTrue(ws3.Drawings.Count < 3);
            ws3.Drawings[1].Copy(ws3, 1, 15, 0, 0);
            Assert.AreEqual(3, ws3.Drawings.Count);
            SaveAndCleanup(p);
        }

        //Copy Range
        [TestMethod]
        public void CopyDrawingsRange()
        {
            using var p = OpenTemplatePackage("CopyDrawings.xlsx");
            var ws3 = p.Workbook.Worksheets[2];
            ws3.Cells["A1:Z50"].Copy(ws3.Cells["AA1:AZ50"]);
            SaveAndCleanup(p);
        }


        private class Item
        {
            public string Name { get; set; }

            public int Value { get; set; }
        }

        [TestMethod]
        public void issue1475()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");


            IEnumerable<Item> _items = new List<Item>() {
                new Item { Name = "Bob", Value = 3 },
                new Item { Name = "Lisa", Value = 8 },
                new Item { Name = "Steve", Value = 5 },
                new Item { Name = "Phil", Value = 2 },
            };

            var range = ws.Cells["A1"].LoadFromCollection(_items, true, TableStyles.Dark1);
            var chart = ws.Drawings.AddLineChart("LineChartWithDroplines", eLineChartType.Line);
            var serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With Droplines";
            chart.AddDropLines();
            chart.DropLine.Border.Width = 2;
            //Set style 12
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle12);


            var cpyWs = package.Workbook.Worksheets.Add("Copy", ws);
            cpyWs.View.TabSelected = false;
            package.SaveAs("C:\\epplusTest\\Testoutput\\i1475.xlsx");
        }
        //i1597
        [TestMethod]
        public void CopyExistingLinkedPicture()
        {
            using (var package = OpenTemplatePackage("i1597.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[0];
                var originalPic = (ExcelPicture)sheet.Drawings[0];

                var newWS = package.Workbook.Worksheets.Copy("Sheet1", "Copy");

                var copiedPic = (ExcelPicture)newWS.Drawings[0];

                Assert.AreEqual(originalPic.LinkedImageRel.TargetUri, copiedPic.LinkedImageRel.TargetUri);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void AddLinkedPictureAndCopy()
        {
            using (var package = OpenPackage("LinkPic.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("emptyWS");
                var uri = GetResourceFile("EPPlus.png").FullName;

                var pic = sheet.Drawings.AddPicture("ImageName", uri, PictureLocation.Link);

                Assert.AreEqual($"file:///{uri}", pic.LinkedImageRel.TargetUri.OriginalString);

                var copiedWs = package.Workbook.Worksheets.Copy("emptyWS", "Copy");
                var picCopied = (ExcelPicture)copiedWs.Drawings[0];
                Assert.AreEqual($"file:///{uri}", picCopied.LinkedImageRel.TargetUri.OriginalString);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void InsertAndLinkPictureAndCopy()
        {
            using (var package = OpenPackage("InsertAndLinkPic.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("emptyWS");
                var uri = GetResourceFile("EPPlus.png").FullName;

                var pic = sheet.Drawings.AddPicture("ImageName", uri, PictureLocation.LinkAndEmbed);

                Assert.AreEqual($"file:///{uri}", pic.LinkedImageRel.TargetUri.OriginalString);

                var copiedWs = package.Workbook.Worksheets.Copy("emptyWS", "Copy");
                var picCopied = (ExcelPicture)copiedWs.Drawings[0];
                Assert.AreEqual($"file:///{uri}", picCopied.LinkedImageRel.TargetUri.OriginalString);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void AddAndCopyImage()
        {
            using (var package = OpenPackage("AddPic.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("emptyWS");
                var uri = GetResourceFile("EPPlus.png").FullName;

                var pic = sheet.Drawings.AddPicture("ImageName", uri);

                var copiedWs = package.Workbook.Worksheets.Copy("emptyWS", "Copy");
                var picCopied = (ExcelPicture)copiedWs.Drawings[0];

                Assert.AreEqual(pic._width, picCopied._width);
                Assert.AreEqual(pic.Size.Width, picCopied.Size.Width);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void AddAndCopyImageWithout100Size()
        {
            using (var package = OpenPackage("AddPic50percent.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("emptyWS");
                var uri = GetResourceFile("EPPlus.png").FullName;

                var pic = sheet.Drawings.AddPicture("ImageName", uri);

                pic.SetSize(50);

                var copiedWs = package.Workbook.Worksheets.Copy("emptyWS", "Copy");
                var picCopied = (ExcelPicture)copiedWs.Drawings[0];

                Assert.AreEqual(pic._width, picCopied._width);
                Assert.AreEqual(pic.Size.Width, picCopied.Size.Width);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ReadAndCopyTwoAnchorImage()
        {
            using (var package = OpenTemplatePackage("SizeCopyTest.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();

                var pic = sheet.Drawings.First();

                sheet.Drawings.ReadPositionsAndSize();

                var copiedWs = package.Workbook.Worksheets.Copy(sheet.Name, "Copy");
                var picCopied = (ExcelPicture)copiedWs.Drawings[0];

                Assert.AreEqual(pic._width, picCopied._width);
                Assert.AreEqual(pic._height, picCopied._height);

                Assert.AreEqual(pic.From.Row, picCopied.From.Row);
                Assert.AreEqual(pic.From.Column, picCopied.From.Column);
                Assert.AreEqual(pic.To.Row, picCopied.To.Row);
                Assert.AreEqual(pic.To.Column, picCopied.To.Column);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void CopyImagesBetweenTwo_Unsaved_Workbooks()
        {
            string sheetName = "AWs";
            string wbName = "CopyPicture_PictureObject";

            using (var src = OpenPackage($"{wbName}1.xlsx", true))
            {
                var wbFirst = src.Workbook;
                var wsFirst = wbFirst.Worksheets.Add(sheetName);

                var pic1 = wsFirst.Drawings.AddPicture("epplusPicture", GetResourceFile("EPPlus.png"));

                using (var target = OpenPackage($"{wbName}2.xlsx", true))
                {
                    var wbSecond = target.Workbook;
                    var wsSecond = wbSecond.Worksheets.Add(sheetName);

                    var pic2 = wsSecond.Drawings.AddPicture("screenshotPicture", GetResourceFile("screenshot.PNG"));
                    pic1.Copy(wsSecond, 1, 1, 0, 0);
                    pic2.Copy(wsFirst, 1, 1, 0, 0);

                    Assert.AreEqual(2, wsSecond.Drawings.Count);
                    Assert.AreEqual(2, wsFirst.Drawings.Count);

                    Assert.AreEqual("epplusPicture", wsSecond.Drawings[1].Name);
                    Assert.AreEqual("screenshotPicture", wsFirst.Drawings[1].Name);

                    //Check that we don't create new rels as image already exists in target
                    Assert.AreEqual(2, wsFirst.Drawings.Part._rels.Count);
                    Assert.AreEqual(2, wsSecond.Drawings.Part._rels.Count);

                    SaveAndCleanup(target);
                }
                SaveAndCleanup(src);
            }
        }

        [TestMethod]
        public void CopyImagesBetweenTwo_UnsavedWorkbooks_NamedRanges()
        {
            string sheetName = "AWs";
            string rangeName = "SomeName";
            string rangeAddress = "A1:G9";
            string wbName = "CopyPicture_NamedRanges";

            using (var src = OpenPackage($"{wbName}1.xlsx", true))
            {
                var wbFirst = src.Workbook;
                var wsFirst = wbFirst.Worksheets.Add(sheetName);

                var pic1 = wsFirst.Drawings.AddPicture("epplusPicture", GetResourceFile("EPPlus.png"));

                wsFirst.Names.AddName(rangeName, wsFirst.Cells[rangeAddress]);

                using (var target = OpenPackage($"{wbName}2.xlsx", true))
                {
                    var wbSecond = target.Workbook;
                    var wsSecond = wbSecond.Worksheets.Add(sheetName);

                    var pic2 = wsSecond.Drawings.AddPicture("screenshotPicture", GetResourceFile("screenshot.PNG"));
                    wsSecond.Names.AddName(rangeName, wsSecond.Cells[rangeAddress]);

                    CopyImagesBetweenWorkbooks(src, target, sheetName, rangeName);

                    //ensure "normal" copying works
                    Assert.AreEqual(2, wsSecond.Drawings.Count);
                    Assert.AreEqual("screenshotPicture", wsSecond.Drawings[0].Name);
                    Assert.AreEqual("epplusPicture", wsSecond.Drawings[1].Name);

                    CopyImagesBetweenWorkbooks(target, src, sheetName, rangeName);

                    //Ensure nothing's changed in the workbook we are copying FROM when copying back
                    Assert.AreEqual(2, wsSecond.Drawings.Count);
                    Assert.AreEqual("screenshotPicture", wsSecond.Drawings[0].Name);
                    Assert.AreEqual("epplusPicture", wsSecond.Drawings[1].Name);

                    //Copies the copy therefore 3
                    Assert.AreEqual(3, wsFirst.Drawings.Count);
                    Assert.AreEqual("epplusPicture", wsFirst.Drawings[0].Name);
                    Assert.AreEqual("screenshotPicture", wsFirst.Drawings[1].Name);
                    Assert.AreEqual("epplusPicture1", wsFirst.Drawings[2].Name);

                    //Check that we don't create new rels as image already exists in target
                    Assert.AreEqual(2, wsFirst.Drawings.Part._rels.Count);
                    Assert.AreEqual(2, wsSecond.Drawings.Part._rels.Count);

                    //Ensure no missmatch of ids
                    var relIdOriginal = wsFirst.Drawings[0].As.Picture.GetRelId();
                    var relIdCopiedBack = wsFirst.Drawings[2].As.Picture.GetRelId();

                    Assert.AreEqual(relIdOriginal, relIdCopiedBack);
                    Assert.AreEqual("rId1", relIdOriginal);
                    Assert.AreEqual("rId2", wsFirst.Drawings[1].As.Picture.GetRelId());

                    SaveAndCleanup(target);
                }
                SaveAndCleanup(src);
            }
        }

        [TestMethod]
        public void CopyImagesBetweenTwo_Saved_Workbooks_NamedRanges()
        {
            string sheetName = "AWs";
            string rangeName = "SomeName";
            string rangeAddress = "A1:G9";
            string wbName = "CopyPictureRead_NamedRanges";

            //Create the workbooks
            using (var src = OpenPackage($"{wbName}1.xlsx", true))
            {
                var wbFirst = src.Workbook;
                var wsFirst = wbFirst.Worksheets.Add(sheetName);

                var pic1 = wsFirst.Drawings.AddPicture("epplusPicture", GetResourceFile("EPPlus.png"));

                wsFirst.Names.AddName(rangeName, wsFirst.Cells[rangeAddress]);

                SaveAndCleanup(src);
            }
            using (var target = OpenPackage($"{wbName}2.xlsx", true))
            {
                var wbSecond = target.Workbook;
                var wsSecond = wbSecond.Worksheets.Add(sheetName);

                var pic2 = wsSecond.Drawings.AddPicture("screenshotPicture", GetResourceFile("screenshot.PNG"));
                wsSecond.Names.AddName(rangeName, wsSecond.Cells[rangeAddress]);

                SaveAndCleanup(target);
            }

            //Read and copy images between the workbooks
            using (var src = OpenPackage($"{wbName}1.xlsx"))
            {
                var wbFirst = src.Workbook;
                var wsFirst = wbFirst.Worksheets.First();

                using (var target = OpenPackage($"{wbName}2.xlsx"))
                {
                    var wbSecond = target.Workbook;
                    var wsSecond = wbSecond.Worksheets.First();

                    Assert.AreEqual(1, wsSecond.Drawings.Count);
                    Assert.AreEqual(1, wsFirst.Drawings.Count);

                    CopyImagesBetweenWorkbooks(src, target, sheetName, rangeName);
                    CopyImagesBetweenWorkbooks(target, src, sheetName, rangeName);

                    Assert.AreEqual(2, wsSecond.Drawings.Count);
                    //Copies the copy therefore 3
                    Assert.AreEqual(3, wsFirst.Drawings.Count);

                    Assert.AreEqual("epplusPicture", wsSecond.Drawings[1].Name);
                    Assert.AreEqual("screenshotPicture", wsFirst.Drawings[1].Name);

                    //Assert.AreEqual("1", wsFirst.Drawings[2].As.Picture.Part._rels["0"].Id);

                    var outputName = GetOutputFile("", "copy_" + target.File.Name).FullName;
                    target.SaveAs(outputName);
                }
                var outputName2 = GetOutputFile("", "copy_" + src.File.Name).FullName;
                src.SaveAs(outputName2);
            }
        }

        [TestMethod]
        public void CopyImages_Read_EnsureWorkaround()
        {
            using (var src = OpenTemplatePackage("ImageInRange1.xlsx"))
            {
                var wb = src.Workbook;
                using (var target = OpenTemplatePackage("ImageInRange2.xlsx"))
                {
                    var wb2 = target.Workbook;

                    CopyNamedRangeToTargetSheetWB(src, target, "sheet1", "SomeName");

                    var outputName = GetOutputFile("", "alt_" + target.File.Name).FullName;
                    target.SaveAs(outputName);
                }
                var outputName2 = GetOutputFile("", "alt_" + src.File.Name).FullName;
                src.SaveAs(outputName2);
            }
        }

        [TestMethod]
        public void CopyImages_Read()
        {
            using (var src = OpenTemplatePackage("ImageInRange1.xlsx"))
            {
                var wb = src.Workbook;
                using (var target = OpenTemplatePackage("ImageInRange2.xlsx"))
                {
                    var wb2 = target.Workbook;

                    CopyImagesBetweenWorkbooks(src, target, "sheet1", "SomeName");

                    SaveAndCleanup(target);
                }
                SaveAndCleanup(src);
            }
        }

        //From i1841 / s814
        public void CopyImagesBetweenWorkbooks(ExcelPackage src, ExcelPackage target, string sheetName, string namedRange)
        {
            var range = src.Workbook.Worksheets[sheetName].Names[namedRange];
            var targetRange = target.Workbook.Worksheets[sheetName].Names[namedRange];
            range.Copy(targetRange);
        }

        public void CopyNamedRangeToTargetSheetWB(ExcelPackage src, ExcelPackage target, string sheetName, string namedRange)
        {
            var srcWs = src.Workbook.Worksheets[sheetName];
            var tmpWs = target.Workbook.Worksheets.Add("tmpCopy", srcWs);

            var range = tmpWs.Names[namedRange];
            var targetRange = target.Workbook.Worksheets[sheetName].Names[namedRange];
            range.Copy(targetRange);

            target.Workbook.Worksheets.Delete(tmpWs);
        }
    }
}
