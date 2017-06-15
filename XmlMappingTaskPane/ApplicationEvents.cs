//Copyright (c) Microsoft Corporation.  All rights reserved.
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace XmlMappingTaskPane
{
    class ApplicationEvents
    {
        private bool m_fLastSdiMode;
        private int m_intLastDocumentCount; 
        private int m_intLastWindowCount;

        private RibbonToggleButton m_buttonMapping;
        private IDictionary<Word.Window, CustomTaskPane> m_dictTaskPanes;

        private Word.ApplicationEvents4_DocumentChangeEventHandler m_ehDocumentChange;
        private Word.ApplicationEvents4_WindowActivateEventHandler m_ehWindowActivate;
        private Word.Window m_wdwinLastWindow;

        internal ApplicationEvents(RibbonToggleButton button, IDictionary<Word.Window, CustomTaskPane> dictTaskPanes)
        {
            //store the application and Ribbon objects
            m_buttonMapping = button;
            m_dictTaskPanes = dictTaskPanes;

            //store MDI/SDI state
            m_fLastSdiMode = Globals.ThisAddIn.Application.ShowWindowsInTaskbar;

            //capture the necessary app events
            m_ehDocumentChange = new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(app_DocumentChange);
            m_ehWindowActivate = new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(app_WindowActivate);
            Globals.ThisAddIn.Application.DocumentChange += m_ehDocumentChange;
            Globals.ThisAddIn.Application.WindowActivate += m_ehWindowActivate;
        }

        /// <summary>
        /// Handle Word's DocumentChange event (a new document has focus).
        /// </summary>
        private void app_DocumentChange()
        {
            //refresh the ribbon button
            UpdateButtonState();

            int intNewDocCount = Globals.ThisAddIn.Application.Documents.Count;

            if (intNewDocCount == 0 && m_intLastDocumentCount == 1)
            {
                //check if we hit the fishbowl
                //delete the current CTP (it should always be the last one)
                Debug.Assert(Globals.ThisAddIn.CustomTaskPanes.Count <= 1, "why are there undeleted CTPs around?");
                if (Globals.ThisAddIn.CustomTaskPanes.Count == 1)
                {
                    Globals.ThisAddIn.CustomTaskPanes.RemoveAt(0);
                }  
            }

            //in MDI, need to update the task pane to show the content for the new document
            if (!Globals.ThisAddIn.Application.ShowWindowsInTaskbar && Globals.ThisAddIn.Application.Documents.Count > 0 && Globals.ThisAddIn.CustomTaskPanes.Count > 0) //we're in MDI and there are documents and task panes created
            {
                ((Controls.ControlMain)Globals.ThisAddIn.CustomTaskPanes[0].Control).RefreshControls(Controls.ControlMain.ChangeReason.DocumentChanged, null, null, null, null, null);
            }

            //save the new document count
            m_intLastDocumentCount = intNewDocCount;
        }

        /// <summary>
        /// Handle Word's WindowActivate event (a new window has focus).
        /// </summary>
        /// <param name="wddoc">A Document object specifying the document that has focus.</param>
        /// <param name="wdwin">A Window object specifying the window that has focus.</param>
        private void app_WindowActivate(Word.Document wddoc, Word.Window wdwin)
        {            
            int intNewWindowCount = Globals.ThisAddIn.Application.Windows.Count;

            //check if we lost a window, if so, clean up its CustomTaskPane
            if (intNewWindowCount - m_intLastWindowCount == -1)
            {
                //a window was closed - remove any lingering CTP from the collection
                if (m_dictTaskPanes.ContainsKey(m_wdwinLastWindow))
                {
                    Globals.ThisAddIn.CustomTaskPanes.Remove(m_dictTaskPanes[m_wdwinLastWindow]);
                    m_dictTaskPanes.Remove(m_wdwinLastWindow);
                }
            }

            //check for SDI<-->MDI changes
            CheckForMdiSdiSwitch();

            //store new window and count
            m_intLastWindowCount = intNewWindowCount;
            m_wdwinLastWindow = wdwin;
        }

        /// <summary>
        /// Update the state of our button on the Ribbon.
        /// </summary>
        private void UpdateButtonState()
        {
            //check for an MDI<-->SDI switch
            CheckForMdiSdiSwitch();

            try
            {
                //only leave it off in the fishbowl
                if (Globals.ThisAddIn.Application.Documents.Count == 0)
                {
                    m_buttonMapping.Checked = false;
                    m_buttonMapping.Enabled = false;
                    return;
                }
                else
                    m_buttonMapping.Enabled = true;
            }
            catch (COMException ex)
            {
                Debug.Fail(ex.Source, ex.Message);
                m_buttonMapping.Enabled = false;
                return;
            }

            //check pressed state
            if (m_fLastSdiMode)
            {
                //get the ctp for this window (or null if there's not one)
                CustomTaskPane ctpPaneForThisWindow = null;
                try
                {
                    Globals.ThisAddIn.TaskPaneList.TryGetValue(Globals.ThisAddIn.Application.ActiveWindow, out ctpPaneForThisWindow);
                }
                catch (COMException ex)
                {
                    Debug.Fail("Failed to get CTP:" + ex.Message);
                }

                //if it's not built, don't check
                if (ctpPaneForThisWindow != null)
                {
                    //if it's visible, down
                    if (ctpPaneForThisWindow.Visible == true)
                        m_buttonMapping.Checked = true;
                    else
                        m_buttonMapping.Checked = false;
                }
                else
                {
                    m_buttonMapping.Checked = false;
                }
            }
            else
            {
                if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
                {
                    Debug.Assert(Globals.ThisAddIn.CustomTaskPanes.Count == 1, "why are there multiple CTPs?");

                    if (Globals.ThisAddIn.CustomTaskPanes[0].Visible)
                        m_buttonMapping.Checked = true;
                    else
                        m_buttonMapping.Checked = false;
                }
            }
        }

        /// <summary>
        /// Check if the application has moved from MDI mode to SDI mode (or vice versa).
        /// </summary>
        private void CheckForMdiSdiSwitch()
        {
            //check if we changed
            if (Globals.ThisAddIn.Application.ShowWindowsInTaskbar != m_fLastSdiMode)
            {
                if (m_fLastSdiMode)
                {
                    //going to MDI
                    //the CTP associated with the active window is the only one we need to keep
                    for (int i = Globals.ThisAddIn.CustomTaskPanes.Count - 1; i >= 0; i--)
                    {
                        try
                        {
                            if (Globals.ThisAddIn.CustomTaskPanes[i].Window != Globals.ThisAddIn.Application.ActiveWindow)
                                Globals.ThisAddIn.CustomTaskPanes.RemoveAt(i);
                        }
                        catch (COMException)
                        {
                            //the task pane was disposed by Office, so just remove it from our collection
                            Globals.ThisAddIn.CustomTaskPanes.RemoveAt(i);
                        }
                    }

                    //clear all SDI window references
                    m_dictTaskPanes.Clear();
                }
                else
                {
                    //going to SDI
                    //if it exists, update the task pane & add it to the list
                    if (Globals.ThisAddIn.CustomTaskPanes.Count == 1)
                    {
                        m_dictTaskPanes.Add((Word.Window)Globals.ThisAddIn.CustomTaskPanes[0].Window, Globals.ThisAddIn.CustomTaskPanes[0]);
                    }
                }

                //switch internal state
                m_fLastSdiMode = !m_fLastSdiMode;

                //update the ribbon
                UpdateButtonState();
            }
        }

        private void ctp_VisibleChange(object ctp, System.EventArgs eventArgs)
        {
            //tell the ribbon to refresh
            UpdateButtonState();
        }

        /// <summary>
        /// Connect the VisibleChanged event for an instance of our task pane.
        /// </summary>
        /// <param name="ctp">A CustomTaskPane specifying the new instance of the task pane.</param>
        public void ConnectTaskPaneEvents(CustomTaskPane ctp)
        {
            ctp.VisibleChanged += new System.EventHandler(ctp_VisibleChange);
        }
    }
}
