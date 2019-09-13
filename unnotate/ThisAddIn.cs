using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

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

        internal static void ToggleShowHideObjects()
        {
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            foreach (PowerPoint.Slide slide in slides)
            {
                Debug.WriteLine("SlideIndex: " + slide.SlideIndex);
                PowerPoint.Slide currSlide = slides[slide.SlideIndex];
                foreach (PowerPoint.Shape shape in currSlide.Shapes)
                {
                    Debug.WriteLine("ShapeID: " + shape.Id);
                    if (shape.Type.Equals(Office.MsoShapeType.msoTextBox))
                    {
                        //Debug.WriteLine("TextBox Color: " + shape.TextFrame.TextRange.Font.Color.RGB);
                        if (shape.TextFrame.TextRange.Font.Color.RGB.Equals(9109675))
                        {
                            shape.Visible = Office.MsoTriState.msoTriStateToggle;
                        }
                    }
                    else if (shape.Type.Equals(Office.MsoShapeType.msoInkComment))
                    {
                        Debug.WriteLine("Ink RGB: " + shape.Line.ForeColor.RGB);
                        if (shape.Line.ForeColor.RGB.Equals(9109675))
                        {
                            shape.Visible = Office.MsoTriState.msoTriStateToggle;
                        }
                    }
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
