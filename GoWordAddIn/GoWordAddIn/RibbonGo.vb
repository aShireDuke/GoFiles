' Created 20150401 By Andrea Dukeshire
' Custom ribbon for GoFile interaction including custom save dialogs
' and common office requirements (print current page, etc)
Imports Microsoft.Office.Tools.Ribbon

Public Class RibbonGo

    ' Location of zPrecedents
    Const WINDOWS_7_PRECEDENT_LOCATION = "c:\Users\Public\Documents\"
    Const PRECEDENT_FOLDER = "zPrecedents"

    Private Sub RibbonGo_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        System.Windows.Forms.MessageBox.Show("Hello you!")

        Dim valReturn As Integer
        Dim suggestName As String
        Dim dlgSaveAs As Microsoft.Office.Core.FileDialog

        ' Hard code suggested name for now...
        suggestName = "Smith-FGo"

        ' Make the file dialog visible to the user 
        dlgSaveAs = Globals.ThisAddIn.Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogSaveAs)

        With dlgSaveAs
            .InitialFileName = suggestName
            .Title = "SaveAs Dialog for SaveAsGoForm"
            .ButtonName = "Save"
            valReturn = .Show()
        End With

        ' If press save then Action = 1, if press cancel then action = 0
        ' Might want to do something if they don't continue...
        If (valReturn = -1) Then
            dlgSaveAs.Execute()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        GetXmlFromResource()
    End Sub

    ' VB implementation of same function in c# in GoWordDoc project
    ' From "Walkthrough: Binding content controls to custom XML parts"
    Private Function GetXmlFromResource() As String

        ' Get filename
        Dim currentFilename As String
        Dim xmlFileName As String
        currentFilename = Globals.ThisAddIn.Application.ActiveDocument.Name
        xmlFileName = currentFilename + ".GoData.xml"


        Dim asm As System.Reflection.Assembly = _
            System.Reflection.Assembly.GetExecutingAssembly()
        Dim stream1 As System.IO.Stream = asm.GetManifestResourceStream( _
                xmlFileName)
        Dim resourceNames As String() = asm.GetManifestResourceNames()

        '"EmployeeControls.employees.xml")


        Using resourceReader As System.IO.StreamReader = _
            New System.IO.StreamReader(stream1)
            If resourceReader IsNot Nothing Then
                Return resourceReader.ReadToEnd()
            End If
        End Using

        Return Nothing

    End Function

End Class
