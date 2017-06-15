
namespace XmlMappingTaskPane
{
    partial class RibbonMapping : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMapping()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tabDeveloper = this.Factory.CreateRibbonTab();
            this.groupMapping = this.Factory.CreateRibbonGroup();
            this.toggleButtonMapping = this.Factory.CreateRibbonToggleButton();
            this.tabDeveloper.SuspendLayout();
            this.groupMapping.SuspendLayout();
            // 
            // tabDeveloper
            // 
            this.tabDeveloper.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabDeveloper.ControlId.OfficeId = "TabDeveloper";
            this.tabDeveloper.Groups.Add(this.groupMapping);
            this.tabDeveloper.Label = "TabDeveloper";
            this.tabDeveloper.Name = "tabDeveloper";
            // 
            // groupMapping
            // 
            this.groupMapping.DialogLauncher = ribbonDialogLauncherImpl1;
            this.groupMapping.Items.Add(this.toggleButtonMapping);
            this.groupMapping.Label = "Mapping";
            this.groupMapping.Name = "groupMapping";
            this.groupMapping.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupControls");
            this.groupMapping.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.groupMapping_DialogLauncherClick);
            // 
            // toggleButtonMapping
            // 
            this.toggleButtonMapping.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.toggleButtonMapping.Image = global::XmlMappingTaskPane.Properties.Resources.RibbonIcon;
            this.toggleButtonMapping.KeyTip = "M";
            this.toggleButtonMapping.Label = "XML Mapping";
            this.toggleButtonMapping.Name = "toggleButtonMapping";
            this.toggleButtonMapping.ScreenTip = "XML Mapping";
            this.toggleButtonMapping.ShowImage = true;
            this.toggleButtonMapping.SuperTip = "Opens the XML mapping task pane, which allows you to create custom XML data and ma" +
                "p it to content controls within the current document.";
            this.toggleButtonMapping.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonMapping_Click);
            // 
            // RibbonMapping
            // 
            this.Name = "RibbonMapping";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabDeveloper);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMapping_Load);
            this.tabDeveloper.ResumeLayout(false);
            this.tabDeveloper.PerformLayout();
            this.groupMapping.ResumeLayout(false);
            this.groupMapping.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabDeveloper;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMapping;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonMapping;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMapping RibbonMapping
        {
            get { return this.GetRibbon<RibbonMapping>(); }
        }
    }
}
