using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using MessageBox = System.Windows.Forms.MessageBox;

namespace unnotate
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal static void ToggleShowHideObjects(bool show)
        {
            Office.MsoTriState isVisible;
            if (show == true)
            {
                isVisible = Office.MsoTriState.msoTrue;
            } else
            {
                isVisible = Office.MsoTriState.msoFalse;
            }

            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            foreach (PowerPoint.Slide slide in slides)
            {
                //Debug.WriteLine("SlideIndex: " + slide.SlideIndex);
                PowerPoint.Slide currSlide = slides[slide.SlideIndex] ;
                foreach (PowerPoint.Shape shape in currSlide.Shapes)
                {
                    //Debug.WriteLine("ShapeID: " + shape.Id);
                    if (shape.Type.Equals(Office.MsoShapeType.msoTextBox))
                    {
                        //Debug.WriteLine("TextBox Color: " + shape.TextFrame.TextRange.Font.Color.RGB);
                        if (shape.TextFrame.TextRange.Font.Color.RGB.Equals(9109675) || shape.TextFrame.TextRange.Font.Color.RGB.Equals(10498160))
                        {
                            shape.Visible = isVisible;
                        }
                    }
                    else if (shape.Type.Equals(Office.MsoShapeType.msoInkComment))
                    {
                        //Debug.WriteLine("Ink RGB: " + shape.Line.ForeColor.RGB);
                        if (shape.Line.ForeColor.RGB.Equals(9109675) || shape.Line.ForeColor.RGB.Equals(10498160))
                        {
                            shape.Visible = isVisible;
                        }
                    }
                }
            }
        }

        internal static void DrawTextBox()
        {
            int slideIdx = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            PowerPoint.Slide currSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides[slideIdx];
            PowerPoint.TextRange newTextRange = currSlide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 200, 50).TextFrame.TextRange;
            newTextRange.Text = "New TextBox";
            //newTextRange.Text = "";
            newTextRange.Font.Size = 18;
            //newTextRange.Font.Color.RGB = Color.Purple.ToArgb();
            newTextRange.Font.Color.RGB = 9109675;
        }

        internal static void RemoveObjects(Boolean warn)
        {
            DialogResult result;
            if (warn)
            {
                string messageBoxText = "Are you sure to remove all the objects in purple?";
                string caption = "Warning";
                MessageBoxButtons button = MessageBoxButtons.OKCancel;
                MessageBoxIcon icon = MessageBoxIcon.Warning;
                result = MessageBox.Show(messageBoxText, caption, button, icon);
            } else
            {
                result = DialogResult.OK;
            }

            switch (result)
            {
                case DialogResult.OK:
                    PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                    while (CheckObjectsExist())
                    {
                        foreach (PowerPoint.Slide slide in slides)
                        {
                            //Debug.WriteLine("SlideIndex: " + slide.SlideIndex);
                            PowerPoint.Slide currSlide = slides[slide.SlideIndex];
                            foreach (PowerPoint.Shape shape in currSlide.Shapes)
                            {
                                //Debug.WriteLine("ShapeID: " + shape.Id);
                                if (shape.Type.Equals(Office.MsoShapeType.msoTextBox))
                                {
                                    //Debug.WriteLine("TextBox Color: " + shape.TextFrame.TextRange.Font.Color.RGB);
                                    if (shape.TextFrame.TextRange.Font.Color.RGB.Equals(9109675))
                                    {
                                        shape.Delete();
                                    }
                                }
                                else if (shape.Type.Equals(Office.MsoShapeType.msoInkComment))
                                {
                                    //Debug.WriteLine("Ink RGB: " + shape.Line.ForeColor.RGB);
                                    if (shape.Line.ForeColor.RGB.Equals(9109675))
                                    {
                                        shape.Delete();
                                    }
                                }
                            }
                        }
                    }
                    break;
                case DialogResult.Cancel:
                    break;
            }
        }

        internal static Boolean CheckObjectsExist() {
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            foreach (PowerPoint.Slide slide in slides)
            {
                PowerPoint.Slide currSlide = slides[slide.SlideIndex];
                foreach (PowerPoint.Shape shape in currSlide.Shapes)
                {
                    if (shape.Type.Equals(Office.MsoShapeType.msoTextBox))
                    {
                        if (shape.TextFrame.TextRange.Font.Color.RGB.Equals(9109675))
                        {
                            return true;
                        }
                    }
                    else if (shape.Type.Equals(Office.MsoShapeType.msoInkComment))
                    {
                        if (shape.Line.ForeColor.RGB.Equals(9109675))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        internal static void ExportPowerPointAndRemoveObjects()
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PowerPoint Presentation|*.pptx|PowerPoint 97-2003 Presentation|*.ppt";
            saveFileDialog.Title = "Save as an unnotated File";
            saveFileDialog.InitialDirectory = presentation.Path;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = presentation.Name.Replace(".pptx", "") + "-unnotated";
            DialogResult result = saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                switch (result)
                {
                    case DialogResult.OK:
                        presentation.SaveCopyAs(saveFileDialog.FileName);
                        Globals.ThisAddIn.Application.Presentations.Open(saveFileDialog.FileName, WithWindow: Office.MsoTriState.msoTrue);
                        RemoveObjects(false);
                        Globals.ThisAddIn.Application.ActivePresentation.Save();
                        //Globals.ThisAddIn.Application.ActivePresentation.Close();
                        break;
                    case DialogResult.Cancel:
                        break;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
