using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs;
using PowerPointLabs.DataSources;
using PowerPointLabs.DrawingsLab;
using PowerPointLabs.DrawingsLab.TestInterface;
using PowerPointLabs.Utils;
using Test.Util;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Test.UnitTest.DrawingLab
{
    [TestClass]
    public class DrawingLabMultiCloneUnitTest : BaseUnitTest
    {
        private DrawingsLabMain _drawingsLab;
        private StubDrawingLabSelectionService _selection;
        private StubDrawingLabDialogService _dialog;

        protected override string GetTestingSlideName()
        {
            return "DrawingLab\\DrawingLabMultiClone.pptx";
        }

        [TestInitialize]
        public void Init()
        {
            _drawingsLab = new DrawingsLabMain(new DrawingLabData());
            _selection = new StubDrawingLabSelectionService();
            _dialog = new StubDrawingLabDialogService();
            _drawingsLab.SetSelectionService(_selection);
            _drawingsLab.SetDialogService(_dialog);
        }

        [TestCleanup]
        public void CleanUp()
        {

        }
            
        [TestMethod]
        [TestCategory("UT")]
        public void DrawingLabMultiClone()
        {
            _selection.CurrentSlide = PpOperations.SelectSlide(1);
            var triangleA = PpOperations.SelectShape("TriangleA")[1];
            var triangleB = PpOperations.SelectShape("TriangleB")[1];

            _dialog.NumericDialogAnswer = 13;

            // Test: Selecting an 1 shape should fail.
            _selection.SelectedShapesArray = new[] { triangleB };
            _drawingsLab.MultiCloneBetweenTool();
            Assert.AreEqual(TextCollection.DrawingsLabSelectTwoSetsOfShapes, _dialog.LastMessageBoxMessage);

            // Test: Selecting no shapes should fail.
            _selection.SelectedShapesArray = new Shape[] { };
            _drawingsLab.MultiCloneExtendTool();
            Assert.AreEqual(TextCollection.DrawingsLabSelectTwoSetsOfShapes, _dialog.LastMessageBoxMessage);


            // Test: MultiClone Between (5)
            _selection.SelectedShapesArray = new[] {triangleB, triangleA};
            _dialog.NumericDialogAnswer = 5;
            _drawingsLab.MultiCloneBetweenTool();

            // Test: MultiClone Extend (3)
            _selection.SelectedShapesArray = new[] {triangleA, triangleB};
            _dialog.NumericDialogAnswer = 3;
            _drawingsLab.MultiCloneExtendTool();


            // There should be a total of 10 shapes in the slide.
            var allShapes = PpOperations.SelectAllShapesInSlide().Cast<Shape>().ToList();
            Assert.AreEqual(10, allShapes.Count);
            

            // Test: Selecting an odd number of shapes should fail.
            var threeShapes = allShapes.GetRange(0,3).ToArray();
            _selection.SelectedShapesArray = threeShapes;
            _drawingsLab.MultiCloneExtendTool();
            Assert.AreEqual(TextCollection.DrawingsLabSelectTwoSetsOfShapes, _dialog.LastMessageBoxMessage);
            _drawingsLab.MultiCloneBetweenTool();
            Assert.AreEqual(TextCollection.DrawingsLabSelectTwoSetsOfShapes, _dialog.LastMessageBoxMessage);


            // There should still be 10 shapes in the slide.
            Assert.AreEqual(10, PpOperations.SelectAllShapesInSlide().Cast<Shape>().ToList().Count);


            // Sort by zOrder position to check that the z-orders are correct as well.
            allShapes.Sort((s1, s2) => s1.ZOrderPosition - s2.ZOrderPosition);

            AssertShapeMatchesParameters(allShapes[0], 20, 24, 65, 69, 0);

            AssertShapeMatchesParameters(allShapes[1], 45, 41, 65, 69, 15);
            AssertShapeMatchesParameters(allShapes[2], 70, 58, 65, 69, 30);
            AssertShapeMatchesParameters(allShapes[3], 95, 75, 65, 69, 45);
            AssertShapeMatchesParameters(allShapes[4], 120, 92, 65, 69, 60);
            AssertShapeMatchesParameters(allShapes[5], 145, 109, 65, 69, 75);

            AssertShapeMatchesParameters(allShapes[6], 170, 126, 65, 69, 90);

            AssertShapeMatchesParameters(allShapes[7], 320, 228, 65, 69, 180);
            AssertShapeMatchesParameters(allShapes[8], 470, 330, 65, 69, 270);
            AssertShapeMatchesParameters(allShapes[9], 620, 432, 65, 69, 0);
        }


        private void AssertShapeMatchesParameters(Shape shape, float left, float top, float width, float height, float rotation)
        {
            /*
            Assert.IsTrue(SlideUtil.IsAlmostSame(left, shape.Left));
            Assert.IsTrue(SlideUtil.IsAlmostSame(top, shape.Top));
            Assert.IsTrue(SlideUtil.IsAlmostSame(width, shape.Width));
            Assert.IsTrue(SlideUtil.IsAlmostSame(height, shape.Height));
            Assert.IsTrue(SlideUtil.IsAlmostSame(rotation, shape.Rotation));
            */

            Assert.AreEqual(left, shape.Left);
            Assert.AreEqual(top, shape.Top);
            Assert.AreEqual(width, shape.Width);
            Assert.AreEqual(height, shape.Height);
            Assert.AreEqual(rotation, shape.Rotation);
        }
    }
}
