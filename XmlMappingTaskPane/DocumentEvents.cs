//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections;
using System.Diagnostics;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace XmlMappingTaskPane
{
    class DocumentEvents
    {
        //task pane control
        private Controls.ControlMain m_cmTaskPane;

        //current part
        private Word.Document m_wddoc;
        private Office.CustomXMLParts m_parts;
        private Office.CustomXMLPart m_currentPart;

        //document/streams/stream event handler storage varaibles
        private ArrayList m_alDocumentEvents = new ArrayList();
        private ArrayList m_alPartsEvents = new ArrayList();
        private ArrayList m_alPartEvents = new ArrayList();

        public DocumentEvents(Controls.ControlMain cm)
        {
            //get the initial part
            m_cmTaskPane = cm;
            m_wddoc = Globals.ThisAddIn.Application.ActiveDocument;
            m_parts = m_wddoc.CustomXMLParts;
            m_currentPart = m_parts[1];

            //hook up event handlers
            SetupEventHandlers();
        }

        /// <summary>
        /// Set up the document-level event handlers.
        /// </summary>
        private void SetupEventHandlers()
        {
            //add the new document level event handlers
            m_alDocumentEvents.Add(new Word.DocumentEvents2_ContentControlOnEnterEventHandler(doc_ContentControlOnEnter));
            m_alDocumentEvents.Add(new Word.DocumentEvents2_ContentControlAfterAddEventHandler(doc_ContentControlAfterAdd));
            m_wddoc.ContentControlOnEnter += (Word.DocumentEvents2_ContentControlOnEnterEventHandler)m_alDocumentEvents[0];
            m_wddoc.ContentControlAfterAdd += (Word.DocumentEvents2_ContentControlAfterAddEventHandler)m_alDocumentEvents[1];

            //set up stream event handlers for this document
            m_alPartsEvents.Add(new Office._CustomXMLPartsEvents_PartAfterAddEventHandler(parts_PartAdd));
            m_alPartsEvents.Add(new Office._CustomXMLPartsEvents_PartBeforeDeleteEventHandler(parts_PartDelete));
            m_alPartsEvents.Add(new Office._CustomXMLPartsEvents_PartAfterLoadEventHandler(parts_PartLoad));
            m_parts.PartAfterAdd += (Office._CustomXMLPartsEvents_PartAfterAddEventHandler)m_alPartsEvents[0];
            m_parts.PartBeforeDelete += (Office._CustomXMLPartsEvents_PartBeforeDeleteEventHandler)m_alPartsEvents[1];
            m_parts.PartAfterLoad += (Office._CustomXMLPartsEvents_PartAfterLoadEventHandler)m_alPartsEvents[2];

            //set up event handlers for the first stream (shown by default)
            m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler(part_NodeAfterDelete));
            m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterInsertEventHandler(part_NodeAfterInsert));
            m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler(part_NodeAfterReplace));
            m_currentPart.NodeAfterDelete += (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
            m_currentPart.NodeAfterInsert += (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
            m_currentPart.NodeAfterReplace += (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];
        }

        /// <summary>
        /// Change the currently active document.
        /// </summary>
        internal void ChangeCurrentDocument()
        {
            Debug.WriteLine("Changing event document to " + Globals.ThisAddIn.Application.ActiveDocument.Name);

            //unhook existing events
            m_wddoc.ContentControlOnEnter -= (Word.DocumentEvents2_ContentControlOnEnterEventHandler)m_alDocumentEvents[0];
            m_wddoc.ContentControlAfterAdd -= (Word.DocumentEvents2_ContentControlAfterAddEventHandler)m_alDocumentEvents[1];
            m_parts.PartAfterAdd -= (Office._CustomXMLPartsEvents_PartAfterAddEventHandler)m_alPartsEvents[0];
            m_parts.PartBeforeDelete -= (Office._CustomXMLPartsEvents_PartBeforeDeleteEventHandler)m_alPartsEvents[1];
            m_parts.PartAfterLoad -= (Office._CustomXMLPartsEvents_PartAfterLoadEventHandler)m_alPartsEvents[2];
            m_currentPart.NodeAfterDelete -= (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
            m_currentPart.NodeAfterInsert -= (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
            m_currentPart.NodeAfterReplace -= (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];

            //release the streams + stream handler references
            m_alDocumentEvents.Clear();
            m_alPartsEvents.Clear();
            m_alPartEvents.Clear();

            //clean up the m_wddoc object (since otherwise the RCW gets disposed out from under it)
            m_wddoc = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            //hook up new objects
            m_wddoc = Globals.ThisAddIn.Application.ActiveDocument;
            m_parts = m_wddoc.CustomXMLParts;
            m_currentPart = m_parts[1];

            SetupEventHandlers();
        }

        /// <summary>
        /// Change the currently active XML part.
        /// </summary>
        /// <param name="cxp">The CustomXMLPart specifying the newly active XML part.</param>
        internal void ChangeCurrentPart(Office.CustomXMLPart cxp)
        {
            if (cxp != null)
            {
                Debug.WriteLine("Changing event stream to " + cxp.Id);

                //unhook the stream event handlers
                m_currentPart.NodeAfterDelete -= (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
                m_currentPart.NodeAfterInsert -= (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
                m_currentPart.NodeAfterReplace -= (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];

                //release the streams + event handler references
                m_currentPart = null;
                m_alPartEvents.Clear();

                //hook up the new stream
                m_currentPart = cxp;

                //set up event handlers on the supplied stream
                m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler(part_NodeAfterDelete));
                m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterInsertEventHandler(part_NodeAfterInsert));
                m_alPartEvents.Add(new Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler(part_NodeAfterReplace));
                m_currentPart.NodeAfterDelete += (Office._CustomXMLPartEvents_NodeAfterDeleteEventHandler)m_alPartEvents[0];
                m_currentPart.NodeAfterInsert += (Office._CustomXMLPartEvents_NodeAfterInsertEventHandler)m_alPartEvents[1];
                m_currentPart.NodeAfterReplace += (Office._CustomXMLPartEvents_NodeAfterReplaceEventHandler)m_alPartEvents[2];
            }
            else
            {
                Debug.Fail("SetCurrentStream received a null stream");
            }
        }

        #region Document-level events

        /// <summary>
        /// Handle Word's OnEnter event for content controls, to set the selection in the pane (if the option is set).
        /// </summary>
        /// <param name="ccEntered">A ContentControl object specifying the control that was entered.</param>
        private void doc_ContentControlOnEnter(Word.ContentControl ccEntered)
        {
            Debug.WriteLine("Document.ContentControlOnEnter fired.");
            if (ccEntered.XMLMapping.IsMapped && ccEntered.XMLMapping.CustomXMLNode != null)
            {
                m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.OnEnter, null, null, null, ccEntered.XMLMapping.CustomXMLNode, null);
            }
        }

        /// <summary>
        /// Handle Word's AfterAdd event for content controls, to set new controls' placeholder text
        /// </summary>
        /// <param name="ccAdded"></param>
        /// <param name="InUndoRedo"></param>
        private void doc_ContentControlAfterAdd(Word.ContentControl ccAdded, bool InUndoRedo)
        {
            if (!InUndoRedo && m_cmTaskPane.RecentDragDrop)
            {
                ccAdded.Application.ScreenUpdating = false;

                //grab the current text in the node (if any)
                string currentText = null;
                if (ccAdded.XMLMapping.IsMapped)
                {
                    currentText = ccAdded.XMLMapping.CustomXMLNode.Text;
                }

                //set the placeholder text (this has the side effect of clearing out the control's contents)
                ccAdded.SetPlaceholderText(null, null, Utilities.GetPlaceholderText(ccAdded.Type));

                //bring back the original text
                if (currentText != null)
                {
                    ccAdded.Range.Text = currentText;
                }

                ccAdded.Application.ScreenUpdating = true;
                ccAdded.Application.ScreenRefresh();
            }
        }

        #endregion

        #region XML part-level events

        private void parts_PartAdd(Office.CustomXMLPart NewStream)
        {
            Debug.WriteLine("Streams.StreamAfterAdd fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.PartAdded, null, null, null, null, null);
        }

        private void parts_PartDelete(Office.CustomXMLPart OldStream)
        {
            Debug.WriteLine("Streams.StreamBeforeDelete fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.PartDeleted, null, null, null, null, OldStream);
        }

        private void parts_PartLoad(Office.CustomXMLPart LoadedStream)
        {
            Debug.WriteLine("Streams.StreamAfterLoad fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.PartLoaded, null, null, null, null, null);
        }

        private void part_NodeAfterDelete(Office.CustomXMLNode mxnDeletedNode, Office.CustomXMLNode mxnDeletedParent, Office.CustomXMLNode mxnDeletedNextSibling, bool bInUndoRedo)
        {
            Debug.WriteLine("Streams.NodeAfterDelete fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.NodeDeleted, mxnDeletedNode, mxnDeletedParent, mxnDeletedNextSibling, null, null);
        }

        private void part_NodeAfterInsert(Office.CustomXMLNode mxnNewNode, bool bInUndoRedo)
        {
            Debug.WriteLine("Streams.NodeAfterInsert fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.NodeAdded, null, null, null, mxnNewNode, null);
        }

        private void part_NodeAfterReplace(Office.CustomXMLNode mxnOldNode, Office.CustomXMLNode mxnNewNode, bool bInUndoRedo)
        {
            Debug.WriteLine("Streams.NodeAfterReplace fired.");
            m_cmTaskPane.RefreshControls(Controls.ControlMain.ChangeReason.NodeReplaced, mxnOldNode, mxnNewNode.ParentNode, mxnNewNode.NextSibling, mxnNewNode, null);
        }

        #endregion

        /// <summary>
        /// Get the currently active XML part. Read-only.
        /// </summary>
        internal Office.CustomXMLPart Part
        {
            get
            {
                return m_currentPart;
            }
        }

        /// <summary>
        /// Get the XML part collection for the current document. Read-only.
        /// </summary>
        internal Office.CustomXMLParts PartCollection
        {
            get
            {
                return m_parts;
            }
        }

        /// <summary>
        /// Get the currently active document. Read-only.
        /// </summary>
        internal Word.Document Document
        {
            get
            {
                return m_wddoc;
            }
        }
    }
}
