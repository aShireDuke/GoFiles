' Module ContentControlHelpers.vb
' Created 20150420 by Andrea Dukeshire
' Module to contain all functions to access/modify content controls within the document

Module ContentControlHelpers

    Public Const SCHEMANAMESPACE = "http://DLOGoFiles.com/namespaces/GoSchema/"

    Sub SaveContentControlsAsXml()

        Dim fileNumber As String = "J-2565"
        Dim clientName As String
        Dim posessionDate As String = "1999-04-01"
        Dim clientTitle As String = "Manager"

        Dim aw As XNamespace = SCHEMANAMESPACE

        fileNumber = getValueContentControl("plainTextContentControl1", 1)
        clientName = getValueContentControl("plainTextContentControl1", 2)
        posessionDate = getValueContentControl("DatePickerContentControl1", 3)
        clientTitle = getValueContentControl("DropDownListContentControl1", 4)

        ' Create namespace, nickname "aw".  By assigining XElement nodes
        ' to this namespace below, we generate a file that has aw as the default 
        ' namespace.  This is expressed by the "xmlns = namspaceName" declaration
        ' at the top of the file.
        Dim xmlDoc As XElement =
            New XElement(aw + "files",
                    New XElement(aw + "file",
                        New XElement(aw + "fileNumber", fileNumber),
                        New XElement(aw + "clientName", clientName),
                        New XElement(aw + "posessionDate", posessionDate),
                        New XElement(aw + "clientTitle", clientTitle)
                ))

        ' Save in main folder
        xmlDoc.Save("C:\\VS2013\\Projects\\GoWordDocVB\\GoWordDocVB\\SmithBlarg.xml")

    End Sub

    ' Function to get value of content control.  For now only tested on type PlainTextContentControls,
    ' will need to expand for more types & do some error checking
    Function getValueContentControl(ByVal ContentControlTitle, ByVal RangeIndex) As String

        Dim CControl As Word.ContentControls
        Dim CControlValue As String
        CControl = Globals.ThisDocument.SelectContentControlsByTitle(ContentControlTitle)
        CControlValue = Globals.ThisDocument.ContentControls(RangeIndex).Range.Text
        Return CControlValue

    End Function

    ' Ensure that the schema is in the library and registered with the document. 
    Private Function CheckSchema() As Boolean
        ' as per MSDN article "XML Schemas and Data in Document-Level Customizations"
        ' https://msdn.microsoft.com/en-us/library/y36t3e16.aspx?cs-save-lang=1&cs-lang=csharp#code-snippet-1

        Const namespaceUri As String = SCHEMANAMESPACE
        Dim namespaceFound As Boolean = False
        Dim namespaceRegistered As Boolean = False

        ' Search for schema in application library -- true for us
        Dim n As Word.XMLNamespace
        For Each n In Globals.ThisDocument.ThisApplication.XMLNamespaces
            If (n.URI = namespaceUri) Then
                namespaceFound = True
            End If
        Next

        If Not namespaceFound Then
            MessageBox.Show("XML Schema is not in library.")
            Return False
        End If

        '' HACK I can't make "Me.XMLSchema..." compile so commented out for now
        '' Search for schema in this document
        'Dim r As Word.XMLSchemaReference
        'For Each r In Me.XMLSchemaReferences
        '    If (r.NamespaceURI = namespaceUri) Then
        '        namespaceRegistered = True
        '    End If
        'Next

        If Not namespaceRegistered Then
            MessageBox.Show("XML Schema is not registered for this document.")
            Return False
        End If

        Return True
    End Function
End Module
