namespace ExcelAddInForMacro
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.btnInjectMacro = this.Factory.CreateRibbonButton();
            this.btnCreateTable = this.Factory.CreateRibbonButton();
            this.btnFormatTable = this.Factory.CreateRibbonButton();
            this.btnRunAll = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Macro Management";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInjectMacro);
            this.group1.Items.Add(this.btnCreateTable);
            this.group1.Items.Add(this.btnFormatTable);
            this.group1.Items.Add(this.btnRunAll);
            this.group1.Label = "Macro Management";
            this.group1.Name = "group1";
            // 
            // btnInjectMacro
            // 
            this.btnInjectMacro.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInjectMacro.Image = global::ExcelAddInForMacro.Properties.Resources.magnet_1;
            this.btnInjectMacro.Label = "Inject Macro";
            this.btnInjectMacro.Name = "btnInjectMacro";
            this.btnInjectMacro.ShowImage = true;
            this.btnInjectMacro.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInjectMacro_Click);
            // 
            // btnCreateTable
            // 
            this.btnCreateTable.Label = "Create Table";
            this.btnCreateTable.Name = "btnCreateTable";
            this.btnCreateTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateTable_Click_1);
            // 
            // btnFormatTable
            // 
            this.btnFormatTable.Label = "Format Table";
            this.btnFormatTable.Name = "btnFormatTable";
            this.btnFormatTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFormatTable_Click);
            // 
            // btnRunAll
            // 
            this.btnRunAll.Label = "Run All";
            this.btnRunAll.Name = "btnRunAll";
            this.btnRunAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRunAll_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFormatTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRunAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInjectMacro;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
