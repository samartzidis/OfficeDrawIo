using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace OfficeDrawIoWord
{
    public partial class Ribbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //System.Windows.Forms.Application.Idle += Application_Idle;

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

            if (shape != null)
                Globals.ThisAddIn.Application.StatusBar = "Draw.io diagram selected. Double-click to edit.";
        }

        //private void Application_Idle(object sender, EventArgs e)
        //{
        //}

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
