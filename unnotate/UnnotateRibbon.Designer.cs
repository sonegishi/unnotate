namespace unnotate
{
    partial class UnnotateRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public UnnotateRibbon()
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
            this.Unnotate = this.Factory.CreateRibbonTab();
            this.objectGroup = this.Factory.CreateRibbonGroup();
            this.showHideToggleButton = this.Factory.CreateRibbonToggleButton();
            this.insertGroup = this.Factory.CreateRibbonGroup();
            this.textBoxButton = this.Factory.CreateRibbonButton();
            this.Unnotate.SuspendLayout();
            this.objectGroup.SuspendLayout();
            this.insertGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // Unnotate
            // 
            this.Unnotate.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Unnotate.Groups.Add(this.insertGroup);
            this.Unnotate.Groups.Add(this.objectGroup);
            this.Unnotate.Label = "Unnotate";
            this.Unnotate.Name = "Unnotate";
            // 
            // objectGroup
            // 
            this.objectGroup.Items.Add(this.showHideToggleButton);
            this.objectGroup.Label = "Objects";
            this.objectGroup.Name = "objectGroup";
            // 
            // showHideToggleButton
            // 
            this.showHideToggleButton.Label = "Show/Hide";
            this.showHideToggleButton.Name = "showHideToggleButton";
            this.showHideToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ToggleButton1_Click);
            // 
            // insertGroup
            // 
            this.insertGroup.Items.Add(this.textBoxButton);
            this.insertGroup.Label = "Insert";
            this.insertGroup.Name = "insertGroup";
            // 
            // textBoxButton
            // 
            this.textBoxButton.Label = "Text Box";
            this.textBoxButton.Name = "textBoxButton";
            this.textBoxButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TextBoxButton_Click);
            // 
            // UnnotateRibbon
            // 
            this.Name = "UnnotateRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.Unnotate);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.UnnotateRibbon_Load);
            this.Unnotate.ResumeLayout(false);
            this.Unnotate.PerformLayout();
            this.objectGroup.ResumeLayout(false);
            this.objectGroup.PerformLayout();
            this.insertGroup.ResumeLayout(false);
            this.insertGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Unnotate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup objectGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton showHideToggleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup insertGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton textBoxButton;
    }

    partial class ThisRibbonCollection
    {
        internal UnnotateRibbon UnnotateRibbon
        {
            get { return this.GetRibbon<UnnotateRibbon>(); }
        }
    }
}
