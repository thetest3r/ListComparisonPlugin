namespace ListProcessingExcelPlugin
{
    partial class ExcelRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelRibbon()
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
            this.ListProcessingTab = this.Factory.CreateRibbonTab();
            this.CommandsGroup = this.Factory.CreateRibbonGroup();
            this.minColEditBox = this.Factory.CreateRibbonEditBox();
            this.maxColEditBox = this.Factory.CreateRibbonEditBox();
            this.helpButton = this.Factory.CreateRibbonButton();
            this.Sheet1ToSheet2 = this.Factory.CreateRibbonButton();
            this.Sheet2ToSheet1 = this.Factory.CreateRibbonButton();
            this.headerRowCheckBox = this.Factory.CreateRibbonCheckBox();
            this.ListProcessingTab.SuspendLayout();
            this.CommandsGroup.SuspendLayout();
            // 
            // ListProcessingTab
            // 
            this.ListProcessingTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ListProcessingTab.Groups.Add(this.CommandsGroup);
            this.ListProcessingTab.Label = "TabAddIns";
            this.ListProcessingTab.Name = "ListProcessingTab";
            // 
            // CommandsGroup
            // 
            this.CommandsGroup.Items.Add(this.minColEditBox);
            this.CommandsGroup.Items.Add(this.maxColEditBox);
            this.CommandsGroup.Items.Add(this.headerRowCheckBox);
            this.CommandsGroup.Items.Add(this.Sheet1ToSheet2);
            this.CommandsGroup.Items.Add(this.Sheet2ToSheet1);
            this.CommandsGroup.Items.Add(this.helpButton);
            this.CommandsGroup.Label = "List Processing";
            this.CommandsGroup.Name = "CommandsGroup";
            // 
            // minColEditBox
            // 
            this.minColEditBox.Label = "Min Column";
            this.minColEditBox.Name = "minColEditBox";
            this.minColEditBox.SuperTip = "The lower boundary in the desired column range";
            this.minColEditBox.Text = null;
            this.minColEditBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.minColEditBox_TextChanged);
            // 
            // maxColEditBox
            // 
            this.maxColEditBox.Label = "Max Column";
            this.maxColEditBox.Name = "maxColEditBox";
            this.maxColEditBox.SuperTip = "The upper boundary in the desired column range";
            this.maxColEditBox.Text = null;
            this.maxColEditBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.maxColEditBox_TextChanged);
            // 
            // helpButton
            // 
            this.helpButton.Label = "Help";
            this.helpButton.Name = "helpButton";
            // 
            // Sheet1ToSheet2
            // 
            this.Sheet1ToSheet2.Label = "Sheet1 to Sheet2";
            this.Sheet1ToSheet2.Name = "Sheet1ToSheet2";
            this.Sheet1ToSheet2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareSheet1_Click);
            // 
            // Sheet2ToSheet1
            // 
            this.Sheet2ToSheet1.Label = "Sheet2 to Sheet1";
            this.Sheet2ToSheet1.Name = "Sheet2ToSheet1";
            this.Sheet2ToSheet1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareSheet2_Click);
            // 
            // headerRowCheckBox
            // 
            this.headerRowCheckBox.Checked = true;
            this.headerRowCheckBox.Label = "Header Row?";
            this.headerRowCheckBox.Name = "headerRowCheckBox";
            // 
            // ExcelRibbon
            // 
            this.Name = "ExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ListProcessingTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.ListProcessingTab.ResumeLayout(false);
            this.ListProcessingTab.PerformLayout();
            this.CommandsGroup.ResumeLayout(false);
            this.CommandsGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ListProcessingTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CommandsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CompareListsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Sheet1ToSheet2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton helpButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Sheet2ToSheet1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox minColEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox maxColEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox headerRowCheckBox;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelRibbon Ribbon1
        {
            get { return this.GetRibbon<ExcelRibbon>(); }
        }
    }
}
