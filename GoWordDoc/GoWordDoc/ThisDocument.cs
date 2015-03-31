using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace GoWordDoc
{    
    public partial class ThisDocument
    {
        [CachedAttribute()]

        // Initialize objects used as loop through adding custom XML parts.
        // (Each addition requires identifying the XMLPartID, an office object, and the namespace prefix.
        public string employeeXMLPartID = string.Empty;
        private Office.CustomXMLPart employeeXMLPart;
        private const string prefix = "xmlns:ns='http://schemas.microsoft.com/vsto/samples'";

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            //// Create debug XML file in bin directory
            //XmlDocGenerator goXmlFile = new XmlDocGenerator("GoXml.xml");
            //goXmlFile.GenerateGoXml("J-54224", "Smith");

            // Bind custom XML to word doc -- as per BindCustomXmlPart.cs
            // Get XML string from the employees.xml file (Resource)
            string xmlData = GetXmlFromResource();

            if (xmlData != null)
            {
                // adds the XML string to a new custom XML part in the document, 
                // and binds the content controls to elements in the custom XML part.
                AddCustomXmlPart(xmlData);
                BindControlsToCustomXmlPart();
            }

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        // This method gets the contents of the XML data file that is embedded as a 
        // resource in the assembly, and returns the contents as an XML string.
        private string GetXmlFromResource()
        {
            // Declare Assembly and a XML stream (our data .xml file)
            System.Reflection.Assembly asm =
                System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.Stream stream1 = asm.GetManifestResourceStream(
                "GoWordDoc.GoXml.xml");

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
        // The AddCustomXmlPart method creates a new custom XML part that contains an XML 
        // string that is passed to the method. To ensure that the custom XML part is only 
        // created once, the method creates the custom XML part only if a custom XML part 
        //with a matching GUID does not already exist in the document. The first time this
        //method is called, it saves the value of the Id property to the employeeXMLPartID
        //string. The value of the employeeXMLPartID string is persisted in the document because 
        //it was declared by using the CachedAttribute attribute.
        private void AddCustomXmlPart(string xmlData)
        {
            if (xmlData != null)
            {
                employeeXMLPart = this.CustomXMLParts.SelectByID(employeeXMLPartID);
                if (employeeXMLPart == null)
                {
                    employeeXMLPart = this.CustomXMLParts.Add(xmlData);
                    employeeXMLPart.NamespaceManager.AddNamespace("ns",
                        @"http://schemas.microsoft.com/vsto/samples");
                    employeeXMLPartID = employeeXMLPart.Id;
                }
            }
        }

        // This method binds each content control to an element in the 
        // custom XML part and sets the date display format of the DatePickerContentControl.
        private void BindControlsToCustomXmlPart()
        {
            string xPathName = "ns:employees/ns:employee/ns:name";
            this.plainTextContentControl1.XMLMapping.SetMapping(xPathName,
                prefix, employeeXMLPart);

            string xPathDate = "ns:employees/ns:employee/ns:hireDate";
            this.datePickerContentControl1.DateDisplayFormat = "MMMM d, yyyy";
            this.datePickerContentControl1.XMLMapping.SetMapping(xPathDate,
                prefix, employeeXMLPart);

            string xPathTitle = "ns:employees/ns:employee/ns:title";
            this.dropDownListContentControl1.XMLMapping.SetMapping(xPathTitle,
                prefix, employeeXMLPart);
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.plainTextContentControl1.Entering += new Microsoft.Office.Tools.Word.ContentControlEnteringEventHandler(this.plainTextContentControl1_Entering);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void plainTextContentControl1_Entering(object sender, ContentControlEnteringEventArgs e)
        {

        }

    }
}
