namespace Avontus.Quantify.ExcelImporter
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tabQuantify = this.Factory.CreateRibbonTab();
            this.tabAvontus = this.Factory.CreateRibbonTab();
            this.gpAvontusQuantify = this.Factory.CreateRibbonGroup();
            this.btnCreateTemplates = this.Factory.CreateRibbonButton();
            this.btnWorkbookValidation = this.Factory.CreateRibbonButton();
            this.btnImportFromExcel = this.Factory.CreateRibbonButton();
            this.gpAvontusQuantifyOptions = this.Factory.CreateRibbonGroup();
            this.btnCheckForUpdates = this.Factory.CreateRibbonButton();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabQuantify.SuspendLayout();
            this.tabAvontus.SuspendLayout();
            this.gpAvontusQuantify.SuspendLayout();
            this.gpAvontusQuantifyOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabQuantify
            // 
            this.tabQuantify.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabQuantify.Label = "TabAddIns";
            this.tabQuantify.Name = "tabQuantify";
            // 
            // tabAvontus
            // 
            this.tabAvontus.Groups.Add(this.gpAvontusQuantify);
            this.tabAvontus.Groups.Add(this.gpAvontusQuantifyOptions);
            this.tabAvontus.Label = "Avontus Quantify Importer";
            this.tabAvontus.Name = "tabAvontus";
            // 
            // gpAvontusQuantify
            // 
            this.gpAvontusQuantify.Items.Add(this.btnCreateTemplates);
            this.gpAvontusQuantify.Items.Add(this.btnWorkbookValidation);
            this.gpAvontusQuantify.Items.Add(this.btnImportFromExcel);
            this.gpAvontusQuantify.Label = "Quantify";
            this.gpAvontusQuantify.Name = "gpAvontusQuantify";
            // 
            // btnCreateTemplates
            // 
            this.btnCreateTemplates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCreateTemplates.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateTemplates.Image")));
            this.btnCreateTemplates.Label = "Create Templates for Import";
            this.btnCreateTemplates.Name = "btnCreateTemplates";
            this.btnCreateTemplates.ShowImage = true;
            this.btnCreateTemplates.SuperTip = "Creates multiple worksheets for data to be entered, which is then used for import" +
    " into Quantify.";
            this.btnCreateTemplates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateTemplates_Click);
            // 
            // btnWorkbookValidation
            // 
            this.btnWorkbookValidation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWorkbookValidation.Image = ((System.Drawing.Image)(resources.GetObject("btnWorkbookValidation.Image")));
            this.btnWorkbookValidation.Label = "Workbook Validation";
            this.btnWorkbookValidation.Name = "btnWorkbookValidation";
            this.btnWorkbookValidation.ShowImage = true;
            this.btnWorkbookValidation.SuperTip = "Validates the import workbook as a whole  based only on it\'s data.";
            this.btnWorkbookValidation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWorkbookValidation_Click);
            // 
            // btnImportFromExcel
            // 
            this.btnImportFromExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportFromExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnImportFromExcel.Image")));
            this.btnImportFromExcel.Label = "Import Current Worksheet";
            this.btnImportFromExcel.Name = "btnImportFromExcel";
            this.btnImportFromExcel.ShowImage = true;
            this.btnImportFromExcel.SuperTip = "Data on the current worksheet is validated and then optionally imported into Quan" +
    "tify.";
            this.btnImportFromExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportFromExcel_Click);
            // 
            // gpAvontusQuantifyOptions
            // 
            this.gpAvontusQuantifyOptions.Items.Add(this.btnCheckForUpdates);
            this.gpAvontusQuantifyOptions.Items.Add(this.btnHelp);
            this.gpAvontusQuantifyOptions.Items.Add(this.btnAbout);
            this.gpAvontusQuantifyOptions.Label = "Options";
            this.gpAvontusQuantifyOptions.Name = "gpAvontusQuantifyOptions";
            // 
            // btnCheckForUpdates
            // 
            this.btnCheckForUpdates.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckForUpdates.Image")));
            this.btnCheckForUpdates.Label = "&Check for Updates...";
            this.btnCheckForUpdates.Name = "btnCheckForUpdates";
            this.btnCheckForUpdates.ShowImage = true;
            this.btnCheckForUpdates.SuperTip = "Check for updates to Quantify Excel Importer";
            this.btnCheckForUpdates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckForUpdates_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
            this.btnHelp.Label = "&Online Help...";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ShowImage = true;
            this.btnHelp.SuperTip = "View help online for the Quantify Excel Importer";
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Image = ((System.Drawing.Image)(resources.GetObject("btnAbout.Image")));
            this.btnAbout.Label = "&About...";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.SuperTip = "View help online for the Quantify Excel Importer";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabQuantify);
            this.Tabs.Add(this.tabAvontus);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AvontusRibbon_Load);
            this.tabQuantify.ResumeLayout(false);
            this.tabQuantify.PerformLayout();
            this.tabAvontus.ResumeLayout(false);
            this.tabAvontus.PerformLayout();
            this.gpAvontusQuantify.ResumeLayout(false);
            this.gpAvontusQuantify.PerformLayout();
            this.gpAvontusQuantifyOptions.ResumeLayout(false);
            this.gpAvontusQuantifyOptions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabQuantify;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tabAvontus;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpAvontusQuantify;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup gpAvontusQuantifyOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckForUpdates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportFromExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateTemplates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWorkbookValidation;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
