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
            Globals.ThisAddIn.SetRibbon(this);

            btnEditDiagram.Enabled = false;
            btnExport.Enabled = false;

            
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
