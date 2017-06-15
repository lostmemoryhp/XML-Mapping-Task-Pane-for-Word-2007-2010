//Copyright (c) Microsoft Corporation.  All rights reserved.
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Tools;
using Microsoft.Win32;
using XmlMappingTaskPane.Controls;
using Word = Microsoft.Office.Interop.Word;

namespace XmlMappingTaskPane
{
    public partial class ThisAddIn
    {
        private ApplicationEvents m_appEvents; // hang on to a reference to all application-level events, so they can't go out of scope
        private IDictionary<Word.Window, CustomTaskPane> m_dicTaskPanes = new Dictionary<Word.Window, CustomTaskPane>(); //key = Word Window object; value = ctp (if any) for that window

        #region Startup/Shutdown

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // start catching application-level events
            m_appEvents = new ApplicationEvents(Globals.Ribbons.RibbonMapping.toggleButtonMapping, m_dicTaskPanes);

            // perform initial registry setup
            try
            {
                if (Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane") == null)
                {
                    Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane", RegistryKeyPermissionCheck.ReadWriteSubTree);

                    using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\XML Mapping Task Pane", true))
                    {
                        rk.SetValue("Options", ControlTreeView.cOptionsShowAttributes + ControlTreeView.cOptionsAutoSelectNode);
                    }

                    //set up the schema library entries
                    int iLocale = int.Parse(Properties.Resources.Locale, CultureInfo.InvariantCulture);
                    SchemaLibrary.SetAlias("http://schemas.openxmlformats.org/package/2006/metadata/core-properties", Properties.Resources.CoreFilePropertiesName, iLocale);
                    SchemaLibrary.SetAlias("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", Properties.Resources.ExtendedFilePropertiesName, iLocale);
                    SchemaLibrary.SetAlias("http://schemas.microsoft.com/office/2006/coverPageProps", Properties.Resources.CoverPagePropertiesName, iLocale);
                }
            }
            catch (System.Security.SecurityException)
            {
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #endregion

        /// <summary>
        /// Start handling the visibility events for a particular taskpane.
        /// </summary>
        /// <param name="ctp">A CustomTaskPane specifying the taskpane whose events we want to handle.</param>
        internal void ConnectTaskPaneEvents(CustomTaskPane ctp)
        {
            m_appEvents.ConnectTaskPaneEvents(ctp);
        }

        /// <summary>
        /// Update all active taskpanes with the new settings.
        /// </summary>
        /// <param name="newOptions">An integer specifying the settings to be applied.</param>
        internal static void UpdateSettings(int newOptions)
        {
            foreach (CustomTaskPane ctp in Globals.ThisAddIn.CustomTaskPanes)
            {
                ((Controls.ControlMain)ctp.Control).RefreshSettings(newOptions);
            }
        }

        /// <summary>
        /// A list of all task panes in the document. Key = Word Window object; Value = CustomTaskPane object
        /// </summary>
        internal IDictionary<Word.Window, CustomTaskPane> TaskPaneList
        {
            get
            {
                return m_dicTaskPanes;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        internal void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
