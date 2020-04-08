namespace AddAttachmentsAddin
{
    partial class RibbonAddAttachments : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonAddAttachments()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonAddAttachments));
            this.tabAddAttachments = this.Factory.CreateRibbonTab();
            this.Attachments = this.Factory.CreateRibbonGroup();
            this.buttonAddAttachments = this.Factory.CreateRibbonButton();
            this.tabAddAttachments.SuspendLayout();
            this.Attachments.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabAddAttachments
            // 
            this.tabAddAttachments.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabAddAttachments.ControlId.OfficeId = "TabSendReceive";
            this.tabAddAttachments.Groups.Add(this.Attachments);
            resources.ApplyResources(this.tabAddAttachments, "tabAddAttachments");
            this.tabAddAttachments.Name = "tabAddAttachments";
            // 
            // Attachments
            // 
            this.Attachments.Items.Add(this.buttonAddAttachments);
            resources.ApplyResources(this.Attachments, "Attachments");
            this.Attachments.Name = "Attachments";
            // 
            // buttonAddAttachments
            // 
            resources.ApplyResources(this.buttonAddAttachments, "buttonAddAttachments");
            this.buttonAddAttachments.Image = global::AddAttachmentsAddin.Properties.Resources.button;
            this.buttonAddAttachments.Name = "buttonAddAttachments";
            this.buttonAddAttachments.ShowImage = true;
            this.buttonAddAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddAttachments_Click);
            // 
            // RibbonAddAttachments
            // 
            this.Name = "RibbonAddAttachments";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabAddAttachments);
            resources.ApplyResources(this, "$this");
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonAddAttachments_Load);
            this.tabAddAttachments.ResumeLayout(false);
            this.tabAddAttachments.PerformLayout();
            this.Attachments.ResumeLayout(false);
            this.Attachments.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAddAttachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Attachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddAttachments;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonAddAttachments RibbonAddAttachments
        {
            get { return this.GetRibbon<RibbonAddAttachments>(); }
        }
    }
}
