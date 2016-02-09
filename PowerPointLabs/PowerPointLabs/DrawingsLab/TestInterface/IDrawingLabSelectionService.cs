using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs.DrawingsLab.TestInterface
{
    public interface IDrawingLabSelectionService
    {
        List<Shape> GetCurrentlySelectedShapes();
        PowerPointSlide GetCurrentSlide();
    }

    public class DrawingLabDefaultSelectionService : IDrawingLabSelectionService
    {
        public List<Shape> GetCurrentlySelectedShapes()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
            {
                return selection.ShapeRange.Cast<Shape>().ToList();
            }
            return new List<Shape>();
        }

        public PowerPointSlide GetCurrentSlide()
        {
            return PowerPointCurrentPresentationInfo.CurrentSlide;
        }
    }

    public class StubDrawingLabSelectionService : IDrawingLabSelectionService
    {
        private List<Shape> shapes;
        private PowerPointSlide slide;

        public List<Shape> GetCurrentlySelectedShapes()
        {
            return shapes;
        }

        public PowerPointSlide GetCurrentSlide()
        {
            return slide;
        }

        public ShapeRange SelectedShapes
        {
            set { shapes = value.Cast<Shape>().ToList(); }
        }

        public Shape[] SelectedShapesArray
        {
            set { shapes = value.ToList(); }
        }

        public Slide CurrentSlide
        {
            set { slide = PowerPointSlide.FromSlideFactory(value); }
        }
    }
}
