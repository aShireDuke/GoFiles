// Created 20150409 by Andrea Dukeshire
// Reference:
// https://vivekcek.wordpress.com/2011/03/10/valide-an-xml-aganist-an-xsd-schema-stored-as-embeded-resource-in-a-dll/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Xml;
using System.Xml.Schema;

namespace GoWordDoc
{
    ///<summary>
    /// Helper class to access schema (embedded resource)
    /// 
    ///</summary>
    public static class GoSchemaAccess
    {
        private const string SCHEMA = "GoSchema.xsd";
        private const string SCHEMANAMESPACE = "http://DLOGoFiles.com/namespaces/GoSchema/";


        ///<summary>
        /// 
        /// 
        ///</summary>

        public static XmlSchemaSet GetSchemaSet()
        {
            // replace this fully qualified name with the schema file
            string schemaPath = typeof(GoSchemaAccess).FullName.Replace("GoSchemaAccess", SCHEMA);
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add(SCHEMANAMESPACE, XmlReader.Create(typeof(GoSchemaAccess).Assembly.GetManifestResourceStream(schemaPath)));
            return schemas;
        }
    }


}
