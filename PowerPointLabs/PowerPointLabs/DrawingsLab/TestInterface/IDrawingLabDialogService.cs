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

    public class StubDrawingLabDialogService : IDrawingLabDialogService
    {
        public int NumericDialogAnswer;
        public string TextDialogAnswer;
        public string LastMessageBoxMessage { get; private set; }
        public string LastMessageBoxCaption { get; private set; }

        public int ShowMultiCloneNumericDialog()
        {
            return NumericDialogAnswer;
        }

        public string ShowInsertTextDialog()
        {
            return TextDialogAnswer;
        }

        public void DisplayMessageBox(string message, string caption)
        {
            LastMessageBoxMessage = message;
            LastMessageBoxCaption = caption;
        }

        public void ClearMessages()
        {
            LastMessageBoxMessage = null;
            LastMessageBoxCaption = null;
        }
    }
}
