using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.DrawingsLab.TestInterface
{
    public interface IDrawingLabDialogService
    {
        int ShowMultiCloneNumericDialog();
        string ShowInsertTextDialog();

        void DisplayMessageBox(string message, string caption);
    }
}
