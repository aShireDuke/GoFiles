README for overall project.

GoFiles consists of three main projects:

#1) GoWordDoc:  Word document with a bound XML file to store client data.  Includes client form & data validation against the schema as per project #2)

#2) GoSchema:  Schema project file (in eclipse) to validate the client data against.  The .xsd schema file generated here gets bound to the goWordDoc as per project #1)

#3) GoWordAddIn:  Word addin to control the go files & various macros used within the office during workflow (print current page, save in specific directory & suggest filename, print envelope, etc.)

====================================================================
INSTALL INSTRUCTIONS:

GoWordDoc:

1)  Run setup.exe from the C:\GoFiles\GoWordDoc\GoWordDoc\publish directory.  This will install .NET and other required dependencies.  
2) Copy the file GoSchema.xsd (located at C:\GoFiles\GoWordDoc\GoWordDoc) to the local machine.
3) Open MS Office Word 2013 on the local machine.  
4) Show the developer tab (file->Options->Customize Ribbon->Developer tab)
5) Add the add-in manually, referencing the GoSchema.xsd you just added in step 2.
