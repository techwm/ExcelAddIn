namespace ExcelAddIn
{
    partial class MSG : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MSG()
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
            this.btnTexterize = this.Factory.CreateRibbonButton();
            this.btnUnTexterize = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnTexterize);
            this.group1.Items.Add(this.btnUnTexterize);
            this.group1.Label = "Texterizer";
            this.group1.Name = "group1";
            // 
            // btnTexterize
            // 
            this.btnTexterize.Image = global::ExcelAddIn.Properties.Resources.T;
            this.btnTexterize.Label = "Texterize";
            this.btnTexterize.Name = "btnTexterize";
            this.btnTexterize.ShowImage = true;
            this.btnTexterize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTexterize_Click);
            // 
            // btnUnTexterize
            // 
            this.btnUnTexterize.Image = global::ExcelAddIn.Properties.Resources.U;
            this.btnUnTexterize.Label = "UnTexterize";
            this.btnUnTexterize.Name = "btnUnTexterize";
            this.btnUnTexterize.ShowImage = true;
            this.btnUnTexterize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnTexterize_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.label1);
            this.group2.Items.Add(this.label2);
            this.group2.Items.Add(this.label3);
            this.group2.Label = "About";
            this.group2.Name = "group2";
            // 
            // label1
            // 
            this.label1.Label = "Created By Eric Koay";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = "Ver 1.0";
            this.label2.Name = "label2";
            // 
            // label3
            // 
            this.label3.Label = "2017 May 28";
            this.label3.Name = "label3";
            // 
            // MSG
            // 
            this.Name = "MSG";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MSG_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTexterize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnTexterize;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
    }

    partial class ThisRibbonCollection
    {
        internal MSG MSG
        {
            get { return this.GetRibbon<MSG>(); }
        }
    }
}
