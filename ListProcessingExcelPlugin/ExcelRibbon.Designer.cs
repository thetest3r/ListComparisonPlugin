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
            this.Sheet1RangeBox = this.Factory.CreateRibbonEditBox();
            this.box1 = this.Factory.CreateRibbonBox();
            this.sheet1HeaderCheckBox = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.sheet2RangeBox = this.Factory.CreateRibbonEditBox();
            this.box2 = this.Factory.CreateRibbonBox();
            this.sheet2HeaderCheckBox = this.Factory.CreateRibbonCheckBox();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.CompareListsButton = this.Factory.CreateRibbonButton();
            this.helpButton = this.Factory.CreateRibbonButton();
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
            this.CommandsGroup.Items.Add(this.Sheet1RangeBox);
            this.CommandsGroup.Items.Add(this.box1);
            this.CommandsGroup.Items.Add(this.sheet1HeaderCheckBox);
            this.CommandsGroup.Items.Add(this.separator1);
            this.CommandsGroup.Items.Add(this.sheet2RangeBox);
            this.CommandsGroup.Items.Add(this.box2);
            this.CommandsGroup.Items.Add(this.sheet2HeaderCheckBox);
            this.CommandsGroup.Items.Add(this.separator2);
            this.CommandsGroup.Items.Add(this.CompareListsButton);
            this.CommandsGroup.Items.Add(this.helpButton);
            this.CommandsGroup.Label = "List Comparison";
            this.CommandsGroup.Name = "CommandsGroup";
            // 
            // Sheet1RangeBox
            // 
            this.Sheet1RangeBox.Label = "Sheet1 Range";
            this.Sheet1RangeBox.MaxLength = 3;
            this.Sheet1RangeBox.Name = "Sheet1RangeBox";
            this.Sheet1RangeBox.SuperTip = "The lower boundary in the desired column range";
            this.Sheet1RangeBox.Text = null;
            this.Sheet1RangeBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheet1Range_TextChanged);
            // 
            // box1
            // 
            this.box1.Name = "box1";
            // 
            // sheet1HeaderCheckBox
            // 
            this.sheet1HeaderCheckBox.Checked = true;
            this.sheet1HeaderCheckBox.Label = "Header Row?";
            this.sheet1HeaderCheckBox.Name = "sheet1HeaderCheckBox";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // sheet2RangeBox
            // 
            this.sheet2RangeBox.Label = "Sheet2 Range";
            this.sheet2RangeBox.Name = "sheet2RangeBox";
            this.sheet2RangeBox.SuperTip = "The upper boundary in the desired column range";
            this.sheet2RangeBox.Text = null;
            this.sheet2RangeBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheet2Range_TextChanged);
            // 
            // box2
            // 
            this.box2.Name = "box2";
            // 
            // sheet2HeaderCheckBox
            // 
            this.sheet2HeaderCheckBox.Checked = true;
            this.sheet2HeaderCheckBox.Label = "Header Row?";
            this.sheet2HeaderCheckBox.Name = "sheet2HeaderCheckBox";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // CompareListsButton
            // 
            this.CompareListsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CompareListsButton.Label = "Compare Lists";
            this.CompareListsButton.Name = "CompareListsButton";
            this.CompareListsButton.ShowImage = true;
            this.CompareListsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareSheetsButton_Click);
            // 
            // helpButton
            // 
            this.helpButton.Label = "Help";
            this.helpButton.Name = "helpButton";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton helpButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Sheet1RangeBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox sheet2RangeBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox sheet1HeaderCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox sheet2HeaderCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelRibbon Ribbon1
        {
            get { return this.GetRibbon<ExcelRibbon>(); }
        }
    }
}
