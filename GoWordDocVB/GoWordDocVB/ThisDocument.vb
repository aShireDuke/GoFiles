Public Class ThisDocument

    Private Sub ThisDocument_Startup() Handles Me.Startup
        MessageBox.Show("XML Schema is not in library.")
    End Sub

    Private Sub ThisDocument_Shutdown() Handles Me.Shutdown
        SaveContentControlsAsXml()
    End Sub

End Class
