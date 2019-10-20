using System;
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
                ThisAddIn.ToggleShowHideObjects(false);
                Globals.Ribbons.UnnotateRibbon.showHideLabel.Label = "Status: Hidden";
            } else
            {
                ThisAddIn.ToggleShowHideObjects(true);
                Globals.Ribbons.UnnotateRibbon.showHideLabel.Label = "Status: Shown";
            }
        }

        private void TextBoxButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.DrawTextBox();
            if (Globals.Ribbons.UnnotateRibbon.showHideToggleButton.Checked)
            {
                ThisAddIn.ToggleShowHideObjects(false);
            } else
            {
                ThisAddIn.ToggleShowHideObjects(true);
            }
        }

        private void RemoveExportButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.ExportPowerPointAndRemoveObjects();
        }

        private void RemoveAllButton_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.RemoveObjects(true);
        }
    }
}
