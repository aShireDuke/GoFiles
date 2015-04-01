Module AutomaticMacros

    Option Explicit 
    ' MODULE: AutomaticMacros.  Macros that are run automatically with certain word commands.
    ' By giving a macro a special name, you can run it automatically when you perform
    ' an operation such as starting Microsoft Word or opening a document. Word recognizes
    ' the following names as automatic macros, or "auto" macros.  Macros that fit this
    ' criteria are included in this module. Also available to the programmer are the functions
    ' autoExec, AutoNew, , AutoOpen, AutoClose, AutoExit, AutoNew.  Search the web for
    ' more details on each of the automatic functions.
    ' REFERENCE: article "auto macros" on msdn office development center.
    '            http://msdn2.microsoft.com/en-us/library/aa211920(office.11).aspx
    '***********************************************************

    Sub autoOpen()
        ' AutoOpen Macro - 10/4/2005 by Grant Dukeshire
        ' DESCRIPTION: This is a macro that is run everytime word opens.  For now,
        '   Opens all GoFiles showing codes (instead of results).  Could be later improved on to
        '   do more upon word startup.  Ideas for this includes

        Dim goPosition As Integer      '   from W00-WGo1
        Dim protection As String
        Dim currentMode As String
        On Error GoTo ErrorHandler
        protection = ActiveDocument.protectionType

        ' Check if ActiveDoc has a "Go" inside it (other than a name like Gorge-etc)
        ' Check that the document is unprotected - ie an old style goFile (not Forms)
        goPosition = InStr(2, ActiveDocument.name, "Go")

        If goPosition > 0 And protection = wdNoProtection Then
            ActiveWindow.View.ShowFieldCodes = True
        Else
            ActiveWindow.View.ShowFieldCodes = False
            If (protection = wdAllowOnlyFormFields Or protection = wdAllowOnlyReading) Then
                ' for any Go form precedent (read only) or any client go (allowFormFields)
                ' have the document cursor go to the very top (as to show title nicely)
                Selection.HomeKey Unit:=wdStory
                currentMode = getMode
                If (currentMode = FUNCTION_ERROR) Then GoTo ErrorHandler
            End If

        End If

        ' Oct 23_2010 Code written to make changing over precedents easier -- change zoom,
        '   view, show merge codes, and show all P, etc formatting markers.
        '   Note by default, it opens to formatting last saved it in.
        ActiveWindow.ActivePane.View.Zoom.Percentage = 100
        ActiveWindow.ActivePane.View.Type = wdNormalView 'wdPrintView

        'Application.Run MacroName:="Normal.General.toggleCodes"
        'With Selection.PageSetup
        '    .DifferentFirstPageHeaderFooter = True
        'End With
        'ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        'With Selection.HeaderFooter.PageNumbers
        '    .RestartNumberingAtSection = True
        '    .StartingNumber = 1
        'End With
        'ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        'ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View.ShowAll

        Exit Sub
ErrorHandler:
        msgBox("Error in autoOpen: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
    End Sub

    Sub autoClose()
        ' AutoClose - Andrea December 12, 2007
        ' DESCRIPTION: This macro is run everytime a document is exited.  This function was written to be
        '   used with the functions in the module "FormUserCustomization", when a user would like to change
        '   the defaults in the go file, without having to do any programming.
        Dim protection As String
        Dim currentMode As String
        Dim status As String
        On Error GoTo ErrorHandler

        protection = ActiveDocument.protectionType
        ' only want to look at modes, etc if it has protection of some type.
        If (protection = wdAllowOnlyFormFields Or protection = wdAllowOnlyReading) Then
            currentMode = getMode
        Else
            Exit Sub
        End If

        If (currentMode = DEFAULT_MODE Or currentMode = HELP_MODE) Then
            If (protection <> wdAllowOnlyFormFields) Then GoTo ErrorHandler
            exitDefaultOrHelpMode(currentMode)
        ElseIf (currentMode = CLIENT_MODE) Then
            If (protection <> wdAllowOnlyFormFields) Then GoTo ErrorHandler
            status = updateAllBookmarks
            If (status = FUNCTION_ERROR) Then GoTo ErrorHandler
        End If

        Exit Sub

ErrorHandler:
        msgBox("Error in autoClose: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
    End Sub

End Module















