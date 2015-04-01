Imports Microsoft.Office.Tools.Ribbon

Public Class RibbonGo

    Private Sub RibbonGo_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        System.Windows.Forms.MessageBox.Show("Hello you!")

        Dim valReturn As Integer

        'make the file dialog visible to the user 
        With Globals.ThisAddIn.Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogSaveAs)
            .InitialFileName = "blaaarg"
            valReturn = .Show()
        End With
        System.Windows.Forms.MessageBox.Show(valReturn)

    End Sub
End Class
