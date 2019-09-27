﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using Microsoft.Office.Tools.Ribbon;

namespace unnotate
{
    public partial class UnnotateRibbon
    {
        private void UnnotateRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ToggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.Ribbons.UnnotateRibbon.showHideToggleButton.Checked)
            {
                //Debug.WriteLine("HERE: Checked");
                ThisAddIn.ToggleShowHideObjects();
            } else
            {
                //Debug.WriteLine("HERE: Unchecked");
                ThisAddIn.ToggleShowHideObjects();
            }
        }

        private void TextBoxButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.DrawTextBox();
        }

        private void DuplicateSlideButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.CurrentSlideDuplicate();
        }

        private void RemoveExportButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.ExportPowerPointAndRemoveObjects();
        }

        private void RemoveAllButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.RemoveObjects();
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.ExportPowerPoint();
        }
    }
}
