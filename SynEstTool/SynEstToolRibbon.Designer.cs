namespace SynEstTool
{
    partial class SynEstToolRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SynEstToolRibbon()
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
            this.synestribbon = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Consolidate = this.Factory.CreateRibbonButton();
            this.BtnStart = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.Btn_PrintEstList = this.Factory.CreateRibbonButton();
            this.Btn_ColMap = this.Factory.CreateRibbonButton();
            this.BtnRevCrit = this.Factory.CreateRibbonButton();
            this.synestribbon.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // synestribbon
            // 
            this.synestribbon.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.synestribbon.Groups.Add(this.group1);
            this.synestribbon.Groups.Add(this.group2);
            this.synestribbon.Label = "SynEst Tool";
            this.synestribbon.Name = "synestribbon";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Consolidate);
            this.group1.Items.Add(this.BtnStart);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // Consolidate
            // 
            this.Consolidate.Label = "Consolidate";
            this.Consolidate.Name = "Consolidate";
            this.Consolidate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Consolidate_Click);
            // 
            // BtnStart
            // 
            this.BtnStart.Label = "test button";
            this.BtnStart.Name = "BtnStart";
            this.BtnStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnStart_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.Btn_PrintEstList);
            this.group2.Items.Add(this.Btn_ColMap);
            this.group2.Items.Add(this.BtnRevCrit);
            this.group2.Label = "Active Estimate";
            this.group2.Name = "group2";
            // 
            // Btn_PrintEstList
            // 
            this.Btn_PrintEstList.Label = "Print Est. List";
            this.Btn_PrintEstList.Name = "Btn_PrintEstList";
            this.Btn_PrintEstList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_PrintEstList_Click);
            // 
            // Btn_ColMap
            // 
            this.Btn_ColMap.Label = "Column Mapping";
            this.Btn_ColMap.Name = "Btn_ColMap";
            this.Btn_ColMap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_ColMap_Click);
            // 
            // BtnRevCrit
            // 
            this.BtnRevCrit.Label = "Review Criteria";
            this.BtnRevCrit.Name = "BtnRevCrit";
            this.BtnRevCrit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRevCrit_Click);
            // 
            // SynEstToolRibbon
            // 
            this.Name = "SynEstToolRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.synestribbon);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SynEstToolRibbon_Load);
            this.synestribbon.ResumeLayout(false);
            this.synestribbon.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab synestribbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Consolidate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_PrintEstList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_ColMap;
        public Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRevCrit;
    }

    partial class ThisRibbonCollection
    {
        internal SynEstToolRibbon SynEstToolRibbon
        {
            get { return this.GetRibbon<SynEstToolRibbon>(); }
        }
    }
}
