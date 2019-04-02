using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OfficeDrawIo
{
    public partial class Ribbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            System.Windows.Forms.Application.Idle += Application_Idle;

            btnEditDiagram.Enabled = false;
            btnExport.Enabled = false;
        }

        private void Application_Idle(object sender, EventArgs e)
        {
            var selectedCtrl = Globals.ThisAddIn.SelectedCtrl;

            btnEditDiagram.Enabled = selectedCtrl != null;
            btnExport.Enabled = selectedCtrl != null;
            btnAddDiagram.Enabled = selectedCtrl == null;
        }

        private void btnAddDiagram_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddDrawIoDiagramOnDocument();
        }

        private void btnEditDiagram_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.EditDrawIoDiagramOnDocument();
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
