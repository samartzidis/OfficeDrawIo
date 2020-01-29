using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OfficeDrawIoPpt
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            btnEditDiagram.Enabled = false;
            btnExport.Enabled = false;

            Globals.ThisAddIn.SelectionChanged += SelectionChanged;
        }

        private void SelectionChanged(object sender, EventArgs e)
        {
            var shape = Globals.ThisAddIn.SelectedShape;

            btnEditDiagram.Enabled = shape != null;
            btnExport.Enabled = shape != null;
            btnAddDiagram.Enabled = shape == null;
        }

        private void btnAddDiagram_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddDiagramOnDocument();
        }

        private void btnEditDiagram_Click(object sender, RibbonControlEventArgs e)
        {
            var shape = Globals.ThisAddIn.SelectedShape;
            Globals.ThisAddIn.EditDiagramShape(shape);
        }

        private void btnExport_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ExportDrawIoDiagram();
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Settings();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.About();
        }
    }
}
