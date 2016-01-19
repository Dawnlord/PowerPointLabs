﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.DrawingsLab
{
    public class DrawingsLabDialogs
    {
        public static int ShowNumericDialog(string text, string caption)
        {
            var prompt = new Form()
            {
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false,
                MaximizeBox = false,
                Width = 160,
                Height = 130,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen,
            };

            var cancel = new Button();
            cancel.Click += (sender, e) => prompt.Close();
            prompt.CancelButton = cancel;

            var textLabel = new Label()
            {
                Top = 10,
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = false,
                Width = prompt.Width
            };

            var textBox = new NumericUpDown() { Left = 20, Top = 40, Width = 120, Height = 80, Text = "5" };
            var confirmation = new Button() { Text = "Ok", Left = 30, Top = 70, Width = 100, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            textBox.Select(0, textBox.Text.Length);

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                int inputValue;
                if (int.TryParse(textBox.Text, out inputValue))
                {
                    return inputValue;
                }
            }
            return -1;
        }

        public static string ShowTextDialog(string caption)
        {
            var prompt = new Form()
            {
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false,
                MaximizeBox = false,
                Width = 200,
                Height = 110,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen,
            };

            var cancel = new Button();
            cancel.Click += (sender, e) => prompt.Close();
            prompt.CancelButton = cancel;

            var textBox = new TextBox() { Left = 20, Top = 20, Width = 160, Height = 80, Text = "" };
            var confirmation = new Button() { Text = "Ok", Left = 50, Top = 50, Width = 100, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.AcceptButton = confirmation;

            textBox.Select(0, textBox.Text.Length);

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                return textBox.Text;
            }
            return null;
        }

        public static int ShowMultiCloneNumericDialog()
        {
            return ShowNumericDialog(TextCollection.DrawingsLabMultiCloneDialogText,
                                     TextCollection.DrawingsLabMultiCloneDialogHeader);
        }

        public static string ShowInsertTextDialog()
        {
            return ShowTextDialog(TextCollection.DrawingsLabSetTextDialogHeader);
        }
    }
}
