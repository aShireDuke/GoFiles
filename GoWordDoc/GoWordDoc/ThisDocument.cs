﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Windows.Forms;



namespace GoWordDoc
{
    /// <summary>
    /// This document class is autogenerated by the VS project to serve as a communication link
    /// between Word and your custom code.  ThisDocument class gives you access to members of the 
    /// Document host item to perform basic tasks in your customization, such as running code when 
    /// the document is opened or closed. MSDN "Document Host" article: https://msdn.microsoft.com/en-us/library/zzf9223t.aspx
    /// This class also reuses code from MSDN tutorial: "Walkthrough: Binding Content Controls to Custom XML Parts"
    /// </summary>
    public partial class ThisDocument
    {

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            //// HACK create debug XML file in main directory.  If want to use it
            //// manually copy over to be the GoData.xml that is embedded in this project
            //XmlDocGenerator goXmlFile = new XmlDocGenerator("GeneratedData.xml");
            //goXmlFile.GenerateGoXml("J-54224", "Smith", "1999-04-01", "Manager");

            // Read embedded xml file, and bind to content controls in the document
            GoCustomXmlPart GoXmlResource = new GoCustomXmlPart("GoWordDoc.GoData.xml");
            string xmlData = GoXmlResource.GetXmlFromEmbeddedResource();

            if (xmlData != null)
            {
                GoXmlResource.AddCustomXmlPart(xmlData);
                GoXmlResource.BindControlsToCustomXmlPart();
            }

            GoXmlResource.CheckSchema();
            //XmlSchemaSet schemas = GoSchemaAccess.GetSchemaSet();

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {

        }



        #region VSTO Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);
        }

        #endregion

    }
}
