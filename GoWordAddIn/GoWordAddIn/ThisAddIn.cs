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
            // display ribbon at runtime
            CreateRibbonExtensibilityObject();

            // Connect the Application_DocumentBeforeSave event handler with the DocumentBeforeSave event.
            this.Application.DocumentBeforeSave +=
                new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        // Added via "Walkthrough: Creating Your First Application-Level Add-in for Word"
        // https://msdn.microsoft.com/en-us/library/cc442946.aspx
        // The new code defines an event handler for the DocumentBeforeSave event, which is raised when 
        // a document is saved. When the user saves a document, the event handler adds new text at the start of the document.
        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            // This code uses an index value of 1 to access the first paragraph in the Paragraphs collection.
            // Although Visual Basic and Visual C# use 0-based arrays, the lower array bounds of most collections 
            // in the Word object model is 1
            Doc.Paragraphs[1].Range.InsertParagraphBefore();
            Doc.Paragraphs[1].Range.Text = "This text was added by using code.";
        }

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
