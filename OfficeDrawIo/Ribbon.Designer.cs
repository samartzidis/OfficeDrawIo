namespace OfficeDrawIo
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnAddDiagram = this.Factory.CreateRibbonButton();
            this.btnEditDiagram = this.Factory.CreateRibbonButton();
            this.btnExport = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnAddDiagram);
            this.group1.Items.Add(this.btnEditDiagram);
            this.group1.Items.Add(this.btnExport);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.btnSettings);
            this.group1.Items.Add(this.btnAbout);
            this.group1.Label = "Draw.io Diagram";
            this.group1.Name = "group1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnAddDiagram
            // 
            this.btnAddDiagram.Image = global::OfficeDrawIo.Properties.Resources.AddControl_16x;
            this.btnAddDiagram.Label = "Add";
            this.btnAddDiagram.Name = "btnAddDiagram";
            this.btnAddDiagram.ShowImage = true;
            this.btnAddDiagram.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddDiagram_Click);
            // 
            // btnEditDiagram
            // 
            this.btnEditDiagram.Image = global::OfficeDrawIo.Properties.Resources.ASX_Edit_blue_16x;
            this.btnEditDiagram.Label = "View/Edit";
            this.btnEditDiagram.Name = "btnEditDiagram";
            this.btnEditDiagram.ShowImage = true;
            this.btnEditDiagram.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditDiagram_Click);
            // 
            // btnExport
            // 
            this.btnExport.Image = global::OfficeDrawIo.Properties.Resources.ExportFile_16x;
            this.btnExport.Label = "Export...";
            this.btnExport.Name = "btnExport";
            this.btnExport.ShowImage = true;
            this.btnExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExport_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.Image = global::OfficeDrawIo.Properties.Resources.Settings_16x;
            this.btnSettings.Label = "Settings...";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Image = global::OfficeDrawIo.Properties.Resources.InformationSymbol_16x;
            this.btnAbout.Label = "About...";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddDiagram;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditDiagram;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon MyRibbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
