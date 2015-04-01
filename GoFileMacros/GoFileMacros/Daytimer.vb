Module Daytimer

                        Option Explicit
    '***********************************************************
    ' MODULE DESCRIPTION: Daytimer.  All functions applicable to the daytimer, and possibly in
    '   the future other general office documents (calc file, personal hours, etc).  This way
    '   the controlling macros do not live locally in any file, and the user will never be
    '   prompted on opening the document to 'enable/disable macros'
    '***********************************************************

    Sub aaaaaopenDaytimer()
        ' openDaytimer - Dec3_07 - Andrea
        ' DESCRIPTION: This function is useful as it lets the user open the daytimer with a shortcut.
        '   The function opens the document and then starts the timer for the automatic close in 2 min
        ' KEYBOARD SHORTCUT: F5

        On Error GoTo ErrorHandler
        ChangeFileOpenDirectory "\\1tbserver\lawoffice\Admin"
        Documents.Open fileName:="Daytimer.doc"
        Application.Run MacroName:="closeAfterTime"
        Exit Sub
ErrorHandler:
        Dim msg As String
        msg = "An error occured! " & "Error number: " & Err.Number & "Error message: " & Err.Description
        msgBox(msg, vbOKOnly + vbCritical)
    End Sub

    Sub aaaacloseAfterTime()
        ' closeAfterTime - Dec3_07 - Andrea
        ' DESCRIPTION:  starts the timer for a document (usually daytimer) to close
        Application.OnTime(Now + TimeValue("00:02:00"), "closeFile")
    End Sub

    Sub aaaacloseFile()
        ' closeFile - december 2007 andrea
        ' DESCRIPTION:  To be used by openDaytimer after 2 minutes has elapsed.  Using the timeout
        '   function was tricky in this case, because if the user had already closed the document,
        '   word would go ahead and try to close it again at the end of two minutes.  This function
        '   determines if the document is still open, and then closes if neccesary.
        Dim fileName As String
        Dim isOpen As Boolean
        On Error GoTo ErrorHandler
        fileName = "Daytimer.doc"
        isOpen = isDocOpen(fileName)
        If isOpen Then
            Application.Documents(fileName).Save()
            Application.Documents(fileName).Close()
        End If
        Exit Sub
ErrorHandler:
        Dim msg As String
        msg = "An error occured! " & "Error number: " & Err.Number & "Error message: " & Err.Description
        msgBox(msg, vbOKOnly + vbCritical)
    End Sub

    Function aaaaaisDocOpen(strDocName As String) As Boolean
        ' isDocOpen - Andrea Dec 2007
        ' DESCRIPTION: returns a boolean of true if the file is open, false if not

        Dim appWord As Object
        Dim wdDoc As Object

        On Error Resume Next
        appWord = GetObject(, "Word.Application")
        If Err <> 0 Then GoTo ErrorHandler
        With appWord
            wdDoc = appWord.Documents(strDocName)
            If Err <> 0 Then GoTo ErrorHandler
        End With
        isDocOpen = True
        Exit Function
ErrorHandler:
        isDocOpen = False

    End Function

End Module

