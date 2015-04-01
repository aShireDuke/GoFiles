// Created 20150328 by Andrea Dukeshire
// Class to generate a Xml file containing client data.  This xml file
// is then bound to the word doc as a embedded resource for the goFile.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace GoWordDoc
{
    public class XmlDocGenerator
    {
        // Properties - get and set for the filename.
        // stores name of Xml file we will generate (do not include relative path for this var)
        public string OutputXmlFilename { get; set;}

        public XmlDocGenerator( string desiredFilename) 
        {
            OutputXmlFilename = desiredFilename;
        }

        // Create most basic GoXml file for starting purposes
        public void GenerateGoXml(string fileNumber, string clientName)
        {
            var xmlDoc =
                new XElement("files",
                  new XElement("file",
                      new XElement("clientName", fileNumber),
                      new XElement("posessionDate", fileNumber),
                      new XElement ("clientTitle", clientName)
                  ));

            xmlDoc.Save(OutputXmlFilename);
        }

        // HACK Initial code taken from Eric White's 'GenerateData.cs'
        // and is a starting point only to framework to write a xml doc
        public void GenerateRandomXml()
	    {
            Random rnd = new Random();
            string[] names = new[] {
            "Bob",
            "Bill",
            "Suzie",
            "Eric",
            "Jim",
            "Cheryl",
            "Andrew",
            "Jack",
            "Celcin",
            "Davies",
            };

            string[] products = new[] {
            "Bike",
            "Unicycle",
            "Car",
            "Plane",
            "Roller skates",
            "Sleigh",
            "Tricycle",
            "Boat",
            };

            int nbrCustomers = 3000;
            int maxLineItems = 6;

            XElement data = new XElement("Customers",
                    Enumerable.Repeat("", nbrCustomers)
                        .Select((s, i) =>
                            new XElement("Customer",
                                new XElement("CustomerID", i + 1),
                                new XElement("Name", names[rnd.Next(names.Length)]),
                                new XElement("HighValueCustomer",
                                    (i % 9 == 0).ToString()),
                                new XElement("Orders",
                                    Enumerable.Repeat("", rnd.Next(maxLineItems) + 1)
                                        .Select((s2, i2) =>
                                            new XElement("Order",
                                                new XElement("ProductDescription", products[rnd.Next(products.Length)]),
                                                new XElement("Quantity", rnd.Next(4) + 1),
                                                new XElement("OrderDate", ((new DateTime(2000, 1, 1)).AddDays(rnd.Next(1000)).ToShortDateString()))
                                            )
                                        )
                                    )
                                )));

            data.Save("DataBlarg.xml");

	    }
    }
}