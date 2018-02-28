namespace ExcelAddInLogging
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
            this.btnCreateInfo = this.Factory.CreateRibbonButton();
            this.btnCreateFatal = this.Factory.CreateRibbonButton();
            this.btnShowLog = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Add-In Log Manager";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCreateInfo);
            this.group1.Items.Add(this.btnCreateFatal);
            this.group1.Items.Add(this.btnShowLog);
            this.group1.Label = "Log Management";
            this.group1.Name = "group1";
            // 
            // btnCreateInfo
            // 
            this.btnCreateInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCreateInfo.Image = global::ExcelAddInLogging.Properties.Resources.info;
            this.btnCreateInfo.Label = "Create Info Log";
            this.btnCreateInfo.Name = "btnCreateInfo";
            this.btnCreateInfo.ShowImage = true;
            this.btnCreateInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateInfo_Click);
            // 
            // btnCreateFatal
            // 
            this.btnCreateFatal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCreateFatal.Image = global::ExcelAddInLogging.Properties.Resources.error;
            this.btnCreateFatal.Label = "Create Fatal Log";
            this.btnCreateFatal.Name = "btnCreateFatal";
            this.btnCreateFatal.ShowImage = true;
            this.btnCreateFatal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateFatal_Click);
            // 
            // btnShowLog
            // 
            this.btnShowLog.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnShowLog.Image = global::ExcelAddInLogging.Properties.Resources.archive_2;
            this.btnShowLog.Label = "Show Log Directory";
            this.btnShowLog.Name = "btnShowLog";
            this.btnShowLog.ShowImage = true;
            this.btnShowLog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowLog_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateFatal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowLog;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
