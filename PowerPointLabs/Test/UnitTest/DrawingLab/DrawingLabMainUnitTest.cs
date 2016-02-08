using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.DataSources;
using PowerPointLabs.DrawingsLab;
using PowerPointLabs.DrawingsLab.TestInterface;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.Utils;
using Test.Util;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Test.UnitTest.DrawingLab
{
    [TestClass]
    public class DrawingLabMainUnitTest : BaseUnitTest
    {
        private DrawingsLabMain _drawingsLab;
        private StubDrawingLabSelectionService _selection;

        protected override string GetTestingSlideName()
        {
            return "DrawingLab\\DrawingLab.pptx";
        }

        [TestInitialize]
        public void Init()
        {
            _drawingsLab = new DrawingsLabMain(new DrawingLabData());
            _selection = new StubDrawingLabSelectionService();
            _drawingsLab.SetSelectionService(_selection);
        }

        [TestCleanup]
        public void CleanUp()
        {
            
        }

        [TestMethod]
        [TestCategory("UT")]
        public void DrawingLabGroupUngroup()
        {
            Assert.AreEqual(0, PpOperations.SelectShapesByPrefix("Group").Count);

            // Group Shapes
            _selection.CurrentSlide = PpOperations.SelectSlide(1);
            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "LeftRectangle", "RightRectangle" });
            _drawingsLab.GroupShapes();

            var groupShapes = PpOperations.SelectShapesByPrefix("Group");
            Assert.AreEqual(1, groupShapes.Count);
            Assert.IsTrue(Graphics.IsAGroup(groupShapes[1]));

            // Ungroup Shapes
            _selection.SelectedShapes = groupShapes;
            _drawingsLab.UngroupShapes();

            var allShapes = PpOperations.SelectShapesByPrefix("");
            Assert.IsTrue(allShapes.Count >= 2);
            foreach (Shape shape in allShapes)
            {
                Assert.IsFalse(Graphics.IsAGroup(shape));
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void DrawingLabToggleArrows()
        {
            var topLine = PpOperations.SelectShape("TopLine")[1];
            var bottomLine = PpOperations.SelectShape("BottomLine")[1];
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, topLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOval, bottomLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.EndArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, bottomLine.Line.EndArrowheadStyle);

            _selection.CurrentSlide = PpOperations.SelectSlide(1);
            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "LeftRectangle", "RightRectangle", "TopLine", "BottomLine" });

            _drawingsLab.ToggleArrowEnd();

            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, topLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOval, bottomLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.EndArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, bottomLine.Line.EndArrowheadStyle);

            _drawingsLab.ToggleArrowStart();

            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOval, bottomLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.EndArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, bottomLine.Line.EndArrowheadStyle);

            _drawingsLab.ToggleArrowEnd();

            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOval, bottomLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, topLine.Line.EndArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, bottomLine.Line.EndArrowheadStyle);

            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "LeftRectangle", "TopLine" });
            _drawingsLab.ToggleArrowEnd();

            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "BottomLine" });
            _drawingsLab.ToggleArrowStart();

            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, bottomLine.Line.BeginArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadOpen, topLine.Line.EndArrowheadStyle);
            Assert.AreEqual(MsoArrowheadStyle.msoArrowheadNone, bottomLine.Line.EndArrowheadStyle);

        }
    }
}
