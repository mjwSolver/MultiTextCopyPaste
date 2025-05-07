namespace MultiTextCopyPaste
{
    partial class MTRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MTRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MTRibbon));
            this.tabMultiText = this.Factory.CreateRibbonTab();
            this.groupCopyPaste = this.Factory.CreateRibbonGroup();
            this.btnMultiCopy = this.Factory.CreateRibbonButton();
            this.btnMultiPaste = this.Factory.CreateRibbonButton();
            this.btnAddToQAT = this.Factory.CreateRibbonButton();
            this.tabMultiText.SuspendLayout();
            this.groupCopyPaste.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMultiText
            // 
            this.tabMultiText.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMultiText.Groups.Add(this.groupCopyPaste);
            this.tabMultiText.Label = "MultiText";
            this.tabMultiText.Name = "tabMultiText";
            // 
            // groupCopyPaste
            // 
            this.groupCopyPaste.Items.Add(this.btnMultiCopy);
            this.groupCopyPaste.Items.Add(this.btnMultiPaste);
            this.groupCopyPaste.Items.Add(this.btnAddToQAT);
            this.groupCopyPaste.Label = "Copy-Paste";
            this.groupCopyPaste.Name = "groupCopyPaste";
            // 
            // btnMultiCopy
            // 
            this.btnMultiCopy.Image = ((System.Drawing.Image)(resources.GetObject("btnMultiCopy.Image")));
            this.btnMultiCopy.ImageName = "Copy_Icon";
            this.btnMultiCopy.Label = "MultiCopy";
            this.btnMultiCopy.Name = "btnMultiCopy";
            this.btnMultiCopy.ScreenTip = "Select multiple textboxes for copying";
            this.btnMultiCopy.ShowImage = true;
            this.btnMultiCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMultiCopy_Click);
            // 
            // btnMultiPaste
            // 
            this.btnMultiPaste.Image = ((System.Drawing.Image)(resources.GetObject("btnMultiPaste.Image")));
            this.btnMultiPaste.Label = "MultiPaste";
            this.btnMultiPaste.Name = "btnMultiPaste";
            this.btnMultiPaste.ScreenTip = "In the same order as copying, select textboxes and paste";
            this.btnMultiPaste.ShowImage = true;
            this.btnMultiPaste.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMultiPaste_Click);
            // 
            // btnAddToQAT
            // 
            this.btnAddToQAT.Label = "";
            this.btnAddToQAT.Name = "btnAddToQAT";
            // 
            // MTRibbon
            // 
            this.Name = "MTRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabMultiText);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MTRibbon_Load);
            this.tabMultiText.ResumeLayout(false);
            this.tabMultiText.PerformLayout();
            this.groupCopyPaste.ResumeLayout(false);
            this.groupCopyPaste.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMultiText;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCopyPaste;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMultiCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMultiPaste;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddToQAT;
    }

    partial class ThisRibbonCollection
    {
        internal MTRibbon MTRibbon
        {
            get { return this.GetRibbon<MTRibbon>(); }
        }
    }
}
