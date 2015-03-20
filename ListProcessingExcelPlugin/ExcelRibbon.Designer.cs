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
            this.ProcessList = this.Factory.CreateRibbonButton();
            this.instructionButton = this.Factory.CreateRibbonButton();
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
            this.CommandsGroup.Items.Add(this.ProcessList);
            this.CommandsGroup.Items.Add(this.instructionButton);
            this.CommandsGroup.Label = "List Processing";
            this.CommandsGroup.Name = "CommandsGroup";
            // 
            // ProcessList
            // 
            this.ProcessList.Label = "Test Button";
            this.ProcessList.Name = "ProcessList";
            this.ProcessList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProcessList_Click);
            // 
            // instructionButton
            // 
            this.instructionButton.Label = "Help";
            this.instructionButton.Name = "instructionButton";
            // 
            // ExcelRibbon
            // 
            this.Name = "ExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ListProcessingTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.ListProcessingTab.ResumeLayout(false);
            this.ListProcessingTab.PerformLayout();
            this.CommandsGroup.ResumeLayout(false);
            this.CommandsGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ListProcessingTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup CommandsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ProcessList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton instructionButton;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelRibbon Ribbon1
        {
            get { return this.GetRibbon<ExcelRibbon>(); }
        }
    }
}
