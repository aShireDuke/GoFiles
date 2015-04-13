// Created 20150409 by Andrea Dukeshire
// Class to embedd a custom XML part in the GoWordDoc, and to link the xml 
// to controlled content fields in the goWordDoc

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Xml.Schema;

namespace GoWordDoc
{
    class GoCustomXmlPart
    {
        // Instructs the Visual Studio Tools for Office runtime to add the specified data object 
        // to the data cache in the document. See method "AddCustomXmlPart" for more details
        [CachedAttribute()]

        // Custom XML part properties
        public string partID = string.Empty;
        private Office.CustomXMLPart CustomPart;
        private string XMLRESOURCE;

        // HACK hardcode namespace for now...
        private const string SCHEMA = "GoSchema.xsd";
        private const string SCHEMANAMESPACE = "http://DLOGoFiles.com/namespaces/GoSchema/";

        public GoCustomXmlPart(string filename)
        {
            XMLRESOURCE = filename;
        }

        ///<summary>
        /// Reads an embedded XML resource file (XMLRESOURCE) from the project into a stream.
        ///</summary>
        ///<returns>
        /// Returns the contents as an XML string.
        /// </returns> 
        public string GetXmlFromEmbeddedResource()
        {
            // Note that this xml file must be added to the project, 
            // then click properties ->  "Build Action" -> "Embedded Resource"
            System.Reflection.Assembly asm =
                System.Reflection.Assembly.GetExecutingAssembly();
            string[] allResourceNames = asm.GetManifestResourceNames();
            System.IO.Stream stream1 = asm.GetManifestResourceStream(
                XMLRESOURCE);

            using (System.IO.StreamReader resourceReader =
                    new System.IO.StreamReader(stream1))
            {
                if (resourceReader != null)
                {
                    return resourceReader.ReadToEnd();
                }
            }

            return null;
        }

        ///<summary>
        ///Creates a new custom XML part that contains an XML 
        ///string that is passed to the method. To ensure that the custom XML part is only 
        ///created once, the method creates the custom XML part only if a custom XML part 
        ///with a matching GUID does not already exist in the document
        ///</summary>
        public void AddCustomXmlPart(string xmlData)
        {
            // The first time this method is called, it saves the value of the Id property to the partID
            //string. The value of the partID string is persisted in the document because 
            //it was declared by using the CachedAttribute attribute.
            if (xmlData != null)
            {
                CustomPart = Globals.ThisDocument.CustomXMLParts.SelectByID(partID);
                if (CustomPart == null)
                {
                    // Add custom XML
                    CustomPart = Globals.ThisDocument.CustomXMLParts.Add(xmlData);

                    // HACK need to fix AddCustomSchemaPart so added there via embedded .xsd 
                    // Add schema by namespace, loads as is a known part of this project
                    CustomPart.NamespaceManager.AddNamespace("ns",
                        @SCHEMANAMESPACE);

                    //XmlSchemaSet schemas = GetSchemaSet();



                    //CustomPart.SchemaCollection.Add(schemas);

                    // replace this fully qualified name with the schema file
                    //string schemaPath = typeof(GoCustomXmlPart).FullName.Replace("GoSchemaAccess", SCHEMA);
                    //XmlSchemaSet schemas = new XmlSchemaSet();
                    //schemas.Add(SCHEMANAMESPACE, XmlReader.Create(typeof(GoSchemaAccess).Assembly.GetManifestResourceStream(schemaPath)));



                    //return xsc;

                    //CustomPart.SchemaCollection.Add(SCHEMANAMESPACE);

                    // Write ID of the custom XML part so we only do it once
                    partID = CustomPart.Id;
                }
            }
        }

        public static XmlSchemaSet GetSchemaSet()
        {
            // replace this fully qualified name with the schema file
            string schemaPath = typeof(GoSchemaAccess).FullName.Replace("GoSchemaAccess", SCHEMA);
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(SCHEMANAMESPACE, XmlReader.Create(typeof(GoSchemaAccess).Assembly.GetManifestResourceStream(schemaPath)));
            return schemas;
        }

        /// <summary>
        /// Binds control to an element in the custom XML part and sets the date display
        /// format of the DatePickerContentControl.
        /// </summary>
        public void BindControlsToCustomXmlPart()
        {
            // Used by BindControlsToCustomXmlPart
            string prefix = "xmlns:ns=\'" + SCHEMANAMESPACE + "\'";

            // Bind each content control in the document to a Xpath query.  
            // Use "this" to refer to the active word doc
            string xPathName = "ns:files/ns:file/ns:clientName";
            Globals.ThisDocument.plainTextContentControl1.XMLMapping.SetMapping(xPathName,
                prefix, CustomPart);

            string xPathDate = "ns:files/ns:file/ns:posessionDate";
            Globals.ThisDocument.datePickerContentControl1.DateDisplayFormat = "MMMM d, yyyy";
            Globals.ThisDocument.datePickerContentControl1.XMLMapping.SetMapping(xPathDate,
                prefix, CustomPart);

            string xPathTitle = "ns:files/ns:file/ns:clientTitle";
            Globals.ThisDocument.dropDownListContentControl1.XMLMapping.SetMapping(xPathTitle,
                prefix, CustomPart);
        }

        /// <summary>
        /// Ensure that the schema is in the library and registered with the document. 
        /// </summary>
        public bool CheckSchema()
        {
            // as per MSDN article "XML Schemas and Data in Document-Level Customizations"
            // https://msdn.microsoft.com/en-us/library/y36t3e16.aspx?cs-save-lang=1&cs-lang=csharp#code-snippet-1

            string namespaceUri = SCHEMANAMESPACE;
            bool namespaceFound = false;
            bool namespaceRegistered = false;

            // Search for schema in application library -- true for us
            foreach (Word.XMLNamespace n in Globals.ThisDocument.Application.XMLNamespaces)
            {
                if (n.URI == namespaceUri)
                {
                    namespaceFound = true;
                }
            }

            if (!namespaceFound)
            {
                MessageBox.Show("XML Schema is not in library.");
                return false;
            }

            // Search for schema in this document
            foreach (Word.XMLSchemaReference r in Globals.ThisDocument.XMLSchemaReferences)
            {
                if (r.NamespaceURI == namespaceUri)
                {
                    namespaceRegistered = true;
                }
            }

            if (!namespaceRegistered)
            {
                MessageBox.Show("XML Schema is not registered for this document: " + namespaceUri);
                return false;
            }

            return true;
        }
    }
}
