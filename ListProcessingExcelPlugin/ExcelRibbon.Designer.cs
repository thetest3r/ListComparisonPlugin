﻿namespace ListProcessingExcelPlugin
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
            this.sheet1DropDown = this.Factory.CreateRibbonDropDown();
            this.refreshButton1 = this.Factory.CreateRibbonButton();
            this.sheet1RangeBox = this.Factory.CreateRibbonEditBox();
            this.sheet1HeaderToggle = this.Factory.CreateRibbonToggleButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.sheet2DropDown = this.Factory.CreateRibbonDropDown();
            this.refreshButton2 = this.Factory.CreateRibbonButton();
            this.sheet2RangeBox = this.Factory.CreateRibbonEditBox();
            this.sheet2HeaderToggle = this.Factory.CreateRibbonToggleButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.capsCheckBox = this.Factory.CreateRibbonCheckBox();
            this.specialCharsCheckBox = this.Factory.CreateRibbonCheckBox();
            this.easterEggButton = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
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
            this.CommandsGroup.Items.Add(this.sheet1DropDown);
            this.CommandsGroup.Items.Add(this.sheet1RangeBox);
            this.CommandsGroup.Items.Add(this.sheet1HeaderToggle);
            this.CommandsGroup.Items.Add(this.separator1);
            this.CommandsGroup.Items.Add(this.sheet2DropDown);
            this.CommandsGroup.Items.Add(this.sheet2RangeBox);
            this.CommandsGroup.Items.Add(this.sheet2HeaderToggle);
            this.CommandsGroup.Items.Add(this.separator2);
            this.CommandsGroup.Items.Add(this.capsCheckBox);
            this.CommandsGroup.Items.Add(this.specialCharsCheckBox);
            this.CommandsGroup.Items.Add(this.easterEggButton);
            this.CommandsGroup.Items.Add(this.separator3);
            this.CommandsGroup.Items.Add(this.CompareListsButton);
            this.CommandsGroup.Items.Add(this.helpButton);
            this.CommandsGroup.Label = "List Comparison";
            this.CommandsGroup.Name = "CommandsGroup";
            // 
            // sheet1DropDown
            // 
            this.sheet1DropDown.Buttons.Add(this.refreshButton1);
            this.sheet1DropDown.Label = "Sheet 1";
            this.sheet1DropDown.Name = "sheet1DropDown";
            this.sheet1DropDown.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sheet1DropDown_ButtonClick);
            // 
            // refreshButton1
            // 
            this.refreshButton1.Label = "Refresh Lists";
            this.refreshButton1.Name = "refreshButton1";
            // 
            // sheet1RangeBox
            // 
            this.sheet1RangeBox.Label = "Columns";
            this.sheet1RangeBox.Name = "sheet1RangeBox";
            this.sheet1RangeBox.SuperTip = "The columns being compared (separated by commas) i.e. a,d,b";
            this.sheet1RangeBox.Text = null;
            this.sheet1RangeBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheet1Range_TextChanged);
            // 
            // sheet1HeaderToggle
            // 
            this.sheet1HeaderToggle.Label = "Contains Header Row (   )";
            this.sheet1HeaderToggle.Name = "sheet1HeaderToggle";
            this.sheet1HeaderToggle.ScreenTip = "Check if first sheet has a header/title row.";
            this.sheet1HeaderToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sheet1HeaderToggle_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // sheet2DropDown
            // 
            this.sheet2DropDown.Buttons.Add(this.refreshButton2);
            this.sheet2DropDown.Label = "Sheet 2";
            this.sheet2DropDown.Name = "sheet2DropDown";
            this.sheet2DropDown.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sheet2DropDown_ButtonClick);
            // 
            // refreshButton2
            // 
            this.refreshButton2.Label = "Refresh Lists";
            this.refreshButton2.Name = "refreshButton2";
            // 
            // sheet2RangeBox
            // 
            this.sheet2RangeBox.Label = "Columns";
            this.sheet2RangeBox.Name = "sheet2RangeBox";
            this.sheet2RangeBox.SuperTip = "The columns being compared (separated by commas) i.e. a,c,b";
            this.sheet2RangeBox.Text = null;
            this.sheet2RangeBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheet2Range_TextChanged);
            // 
            // sheet2HeaderToggle
            // 
            this.sheet2HeaderToggle.Label = "Contains Header Row (   )";
            this.sheet2HeaderToggle.Name = "sheet2HeaderToggle";
            this.sheet2HeaderToggle.ScreenTip = "Check if second sheet has a header/title row.";
            this.sheet2HeaderToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sheet2HeaderToggle_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // capsCheckBox
            // 
            this.capsCheckBox.Label = "Ignore Caps";
            this.capsCheckBox.Name = "capsCheckBox";
            // 
            // specialCharsCheckBox
            // 
            this.specialCharsCheckBox.Label = "Ignore Characters";
            this.specialCharsCheckBox.Name = "specialCharsCheckBox";
            // 
            // easterEggButton
            // 
            this.easterEggButton.Label = "Easter Egg?";
            this.easterEggButton.Name = "easterEggButton";
            this.easterEggButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.easterEggButton_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // CompareListsButton
            // 
            this.CompareListsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CompareListsButton.Image = global::ListProcessingExcelPlugin.Properties.Resources.Icon_39_512;
            this.CompareListsButton.Label = "Compare Lists";
            this.CompareListsButton.Name = "CompareListsButton";
            this.CompareListsButton.ShowImage = true;
            this.CompareListsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareSheetsButton_Click);
            // 
            // helpButton
            // 
            this.helpButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.helpButton.Image = global::ListProcessingExcelPlugin.Properties.Resources.images;
            this.helpButton.Label = "Help";
            this.helpButton.Name = "helpButton";
            this.helpButton.ShowImage = true;
            this.helpButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.helpButton_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox sheet1RangeBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox sheet2RangeBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown sheet1DropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown sheet2DropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton sheet1HeaderToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton sheet2HeaderToggle;
        private Microsoft.Office.Tools.Ribbon.RibbonButton refreshButton1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton refreshButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox capsCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox specialCharsCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton easterEggButton;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelRibbon Ribbon1
        {
            get { return this.GetRibbon<ExcelRibbon>(); }
        }
    }
}
