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
    public class DrawingLabMainUnitTest : BaseUnitTest
    {
        private DrawingsLabMain _drawingsLab;
        private StubDrawingLabSelectionService _selection;
        private StubDrawingLabDialogService _dialog;

        protected override string GetTestingSlideName()
        {
            return "DrawingLab\\DrawingLabMain.pptx";
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
        public void DrawingLabGroupUngroup()
        {
            Assert.AreEqual(0, PpOperations.SelectShapesByPrefix("Group").Count);

            _selection.CurrentSlide = PpOperations.SelectSlide(1);

            // Test : Group requires at least two selected shapes.
            _selection.SelectedShapesArray = new Shape[] { };
            _dialog.ClearMessages();
            _drawingsLab.GroupShapes();
            Assert.AreEqual(TextCollection.DrawingsLabSelectAtLeastTwoShapes, _dialog.LastMessageBoxMessage);

            // Test : Group requires at least two selected shapes.
            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "RightRectangle" });
            _dialog.ClearMessages();
            _drawingsLab.GroupShapes();
            Assert.AreEqual(TextCollection.DrawingsLabSelectAtLeastTwoShapes, _dialog.LastMessageBoxMessage);

            // Group Shapes
            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "LeftRectangle", "RightRectangle" });
            _drawingsLab.GroupShapes();

            var groupShapes = PpOperations.SelectShapesByPrefix("Group");
            Assert.AreEqual(1, groupShapes.Count);
            Assert.IsTrue(Graphics.IsAGroup(groupShapes[1]));

            // Test : Ungroup requires at least one selected shape.
            _selection.SelectedShapesArray = new Shape[] { };
            _dialog.ClearMessages();
            _drawingsLab.UngroupShapes();
            Assert.AreEqual(TextCollection.DrawingsLabSelectAtLeastOneShape, _dialog.LastMessageBoxMessage);


            // Ungroup Shapes
            _selection.SelectedShapes = groupShapes;
            _drawingsLab.UngroupShapes();

            var allShapes = PpOperations.SelectAllShapesInSlide();
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


            // Test : Toggle arrow requires at least one selected shape.
            _selection.SelectedShapesArray = new Shape[] { };
            _dialog.ClearMessages();
            _drawingsLab.ToggleArrowEnd();
            Assert.AreEqual(TextCollection.DrawingsLabSelectAtLeastOneShape, _dialog.LastMessageBoxMessage);
            _dialog.ClearMessages();
            _drawingsLab.ToggleArrowStart();
            Assert.AreEqual(TextCollection.DrawingsLabSelectAtLeastOneShape, _dialog.LastMessageBoxMessage);


            // Test Toggle Arrow
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


        [TestMethod]
        [TestCategory("UT")]
        public void DrawingLabHideUnhide()
        {
            var hiddenCircle = PpOperations.SelectShape("HiddenCircle")[1];
            var leftRectangle = PpOperations.SelectShape("LeftRectangle")[1];
            var topLine = PpOperations.SelectShape("TopLine")[1];


            // Test : Hide requires at least one selected shape.
            _selection.SelectedShapesArray = new Shape[] { };
            _dialog.ClearMessages();
            _drawingsLab.HideTool();
            Assert.AreEqual(TextCollection.DrawingsLabSelectAtLeastOneShape, _dialog.LastMessageBoxMessage);



            Assert.AreEqual(MsoTriState.msoFalse, hiddenCircle.Visible);
            Assert.AreEqual(MsoTriState.msoTrue, leftRectangle.Visible);
            Assert.AreEqual(MsoTriState.msoTrue, topLine.Visible);

            _selection.CurrentSlide = PpOperations.SelectSlide(1);
            _selection.SelectedShapes = PpOperations.SelectShapes(new[] {"RightRectangle", "TopLine"});

            _drawingsLab.HideTool();

            Assert.AreEqual(MsoTriState.msoFalse, hiddenCircle.Visible);
            Assert.AreEqual(MsoTriState.msoTrue, leftRectangle.Visible);
            Assert.AreEqual(MsoTriState.msoFalse, topLine.Visible);

            _drawingsLab.ShowAllTool();

            Assert.AreEqual(MsoTriState.msoTrue, hiddenCircle.Visible);
            Assert.AreEqual(MsoTriState.msoTrue, leftRectangle.Visible);
            Assert.AreEqual(MsoTriState.msoTrue, topLine.Visible);

            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "LeftRectangle" });
            _drawingsLab.HideTool();

            _selection.SelectedShapes = PpOperations.SelectShapes(new[] { "HiddenCircle" });
            _drawingsLab.HideTool();

            Assert.AreEqual(MsoTriState.msoFalse, hiddenCircle.Visible);
            Assert.AreEqual(MsoTriState.msoFalse, leftRectangle.Visible);
            Assert.AreEqual(MsoTriState.msoTrue, topLine.Visible);
        }
    }
}
