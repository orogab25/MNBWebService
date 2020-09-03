namespace MNB
{
    partial class MNBRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MNBRibbon()
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
            this.mnbTab = this.Factory.CreateRibbonTab();
            this.mnbGroup = this.Factory.CreateRibbonGroup();
            this.mnbDownload = this.Factory.CreateRibbonButton();
            this.mnbLog = this.Factory.CreateRibbonButton();
            this.mnbLogSave = this.Factory.CreateRibbonButton();
            this.mnbTab.SuspendLayout();
            this.mnbGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnbTab
            // 
            this.mnbTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.mnbTab.Groups.Add(this.mnbGroup);
            this.mnbTab.Label = "MNB árfolyam";
            this.mnbTab.Name = "mnbTab";
            // 
            // mnbGroup
            // 
            this.mnbGroup.Items.Add(this.mnbDownload);
            this.mnbGroup.Items.Add(this.mnbLog);
            this.mnbGroup.Items.Add(this.mnbLogSave);
            this.mnbGroup.Label = "Orosz Gábor";
            this.mnbGroup.Name = "mnbGroup";
            // 
            // mnbDownload
            // 
            this.mnbDownload.Label = "MNB adatletöltés";
            this.mnbDownload.Name = "mnbDownload";
            this.mnbDownload.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mnbDownload_Click);
            // 
            // mnbLog
            // 
            this.mnbLog.Label = "Log";
            this.mnbLog.Name = "mnbLog";
            this.mnbLog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mnbLog_Click);
            // 
            // mnbLogSave
            // 
            this.mnbLogSave.Label = "Log mentés";
            this.mnbLogSave.Name = "mnbLogSave";
            this.mnbLogSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.mnbLogSave_Click);
            // 
            // MNBRibbon
            // 
            this.Name = "MNBRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.mnbTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MNBRibbon_Load);
            this.mnbTab.ResumeLayout(false);
            this.mnbTab.PerformLayout();
            this.mnbGroup.ResumeLayout(false);
            this.mnbGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab mnbTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup mnbGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mnbDownload;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mnbLog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton mnbLogSave;
    }

    partial class ThisRibbonCollection
    {
        internal MNBRibbon MNBRibbon
        {
            get { return this.GetRibbon<MNBRibbon>(); }
        }
    }
}
