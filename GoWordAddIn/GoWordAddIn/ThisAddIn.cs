// Created 20150328 Andrea Dukeshire
// Basic Excel 2013 Addin project template in visual studio

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace GoWordAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {       
            // Document open event handler - attach our function "InsertTextIntoDoc"
            this.Application.DocumentOpen +=
                new Word.ApplicationEvents4_DocumentOpenEventHandler(InsertTextIntoDoc);

            // New Document event handler - attach our function "InsertTextIntoDoc" 
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument +=
                new Word.ApplicationEvents4_NewDocumentEventHandler(InsertTextIntoDoc);
            
            // display ribbon at runtime
            CreateRibbonExtensibilityObject();

            // Document Before save event handler -- attach our custom function
            this.Application.DocumentBeforeSave +=
                new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // Csutom function to insert some text into the document
        // Added via "Programming Application-Level Add-Ins" MSDN
        // https://msdn.microsoft.com/en-us/library/bb157876.aspx
        private void InsertTextIntoDoc(Microsoft.Office.Interop.Word.Document Doc)
        {
            try
            {
                Word.Range rng = Doc.Range(0, 0);
                rng.Text = "New Text";
                rng.Select();
            }
            catch (Exception ex)
            {
                // Handle exception if for some reason the document is not available.
            }
        }

        // Custom function to add text to document
        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            // Added via "Walkthrough: Creating Your First Application-Level Add-in for Word"
            // https://msdn.microsoft.com/en-us/library/cc442946.aspx
            // Note that in word object model, use 1 based indexing to access first para (not 0)
            Doc.Paragraphs[1].Range.InsertParagraphBefore();
            Doc.Paragraphs[1].Range.Text = "This text was added by using code.";
        }

        // Show our custom ribbon
        protected override Microsoft.Office.Core.IRibbonExtensibility
        CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new RibbonGo() });
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
