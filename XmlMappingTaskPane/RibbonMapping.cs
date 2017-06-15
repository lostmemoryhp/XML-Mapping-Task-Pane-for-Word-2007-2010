//Copyright (c) Microsoft Corporation.  All rights reserved.
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace XmlMappingTaskPane
{
    public partial class RibbonMapping
    {
        private void RibbonMapping_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void toggleButtonMapping_Click(object sender, RibbonControlEventArgs e)
        {
            //get the ctp for this window (or null if there's not one)
            CustomTaskPane ctpPaneForThisWindow = Utilities.FindTaskPaneForCurrentWindow();

            if (toggleButtonMapping.Checked == false)
            {
                Debug.Assert(ctpPaneForThisWindow != null, "how was the ribbon button pressed if there was no control?");

                //it's being unclicked
                if (ctpPaneForThisWindow != null)
                    ctpPaneForThisWindow.Visible = false;
            }
            else
            {
                if (ctpPaneForThisWindow == null)
                {
                    //set the cursor to wait
                    Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorWait;

                    //set up the task pane
                    ctpPaneForThisWindow = Globals.ThisAddIn.CustomTaskPanes.Add(new Controls.ControlMain(), Properties.Resources.TaskPaneName);
                    ctpPaneForThisWindow.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;

                    //store it for later
                    if (Globals.ThisAddIn.Application.ShowWindowsInTaskbar)
                        Globals.ThisAddIn.TaskPaneList.Add(Globals.ThisAddIn.Application.ActiveWindow, ctpPaneForThisWindow);

                    //connect task pane events
                    Globals.ThisAddIn.ConnectTaskPaneEvents(ctpPaneForThisWindow);

                    //get the control we hosted
                    Controls.ControlMain ccm = (Controls.ControlMain)ctpPaneForThisWindow.Control;

                    //hand the eventing class to the control
                    DocumentEvents de = new DocumentEvents(ccm);
                    ccm.EventHandlerAndOnChildren = de;
                    ccm.RefreshControls(Controls.ControlMain.ChangeReason.DocumentChanged, null, null, null, null, null);

                    //show it                            
                    ctpPaneForThisWindow.Visible = true;

                    //reset the cursor
                    Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                }
                else
                {
                    //it's built and being clicked, show it
                    ctpPaneForThisWindow.Visible = true;
                }
            }
        }

        private void groupMapping_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            using (Forms.FormOptions fo = new Forms.FormOptions())
            {
                if (fo.ShowDialog() == DialogResult.OK)
                {
                    ThisAddIn.UpdateSettings(fo.NewOptions);
                }
            }
        }
    }
}
