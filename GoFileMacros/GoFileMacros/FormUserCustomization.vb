Module FormUserCustomization

Option Explicit On
    '***********************************************************
    ' MODULE DESCRIPTION:  FormUserCustomization.  This module contains all of the routines that relate to
    '   individual users or lawoffices wanting to customize the interface to their needs.
    ' Global variables for changingFormDefaults.  isGoFile is either empty or TRUE_STRING
    ' if the document is currently in change default mode, and goFileName contains the name
    ' of the open file.
    '**********************************************************

    ' Will be set to "true" when is a goFile that has entered editing of defaults mode, so that can sense
    ' in autoClose to put the go file back into protected mode.

    Sub toggleDefaultMode()
        ' toggleDefaultMode - Andrea Jan 2008
        ' DESCRIPTION: Enters/Exits changing defaults on Start Goform
        Dim status As String
        On Error GoTo ErrorHandler

        status = toggleDefaultOrHelpMode(DEFAULT_MODE)
        If (status = FUNCTION_ERROR) Then GoTo ErrorHandler

        Exit Sub
ErrorHandler:
        msgBox("Error in toggleDefault Mode")
    End Sub

    Sub toggleHelpMode()
        ' toggleDefaultMode - Andrea Jan 2008
        ' DESCRIPTION: Enters/Exits changing defaults on Start Goform
        Dim status As String
        On Error GoTo ErrorHandler

        status = toggleDefaultOrHelpMode(HELP_MODE)
        If (status = FUNCTION_ERROR) Then GoTo ErrorHandler

        Exit Sub
ErrorHandler:
        msgBox("Error in toggleDefault Mode")

    End Sub

    Function toggleDefaultOrHelpMode(ByVal toggleBetweenMode As String) As String
        ' toggleDefaultOrHelpMode - Andrea, Dec10-2007
        ' DESCRIPTION: Each office will want to have different defaults for every fields, ie uLawyer as Joy vs Erin, etc.
        '   This function puts the user into/out of default changing mode where anything that is inputted to the
        '   goFile is saved as the new default.
        '   calling this macro will toggle between the the argument toggleBetweenMode
        ' Shortcut: Ctrl+Shift+E Reasoning: (kind of a special edit-E of forms)

        Dim protection As String
        Dim currentMode As String

        On Error GoTo ErrorHandler
        protection = ActiveDocument.protectionType
        currentMode = trimWhiteSpace(ActiveDocument.FormFields("goMode").result)

        Select Case currentMode
            Case START_MODE
                If (protection <> wdAllowOnlyReading) Then GoTo ErrorHandler
                enterDefaultOrHelpMode(toggleBetweenMode)
            Case HELP_MODE
                If (protection <> wdAllowOnlyFormFields) Then GoTo ErrorHandler
                ' only want to be able to access help/start mode
                If (toggleBetweenMode = HELP_MODE) Then
                    exitDefaultOrHelpMode(toggleBetweenMode)
                Else
                    GoTo pressedInError
                End If
            Case DEFAULT_MODE
                If (protection <> wdAllowOnlyFormFields) Then GoTo ErrorHandler
                ' only want to be able to access default/start mode
                If (toggleBetweenMode = DEFAULT_MODE) Then
                    exitDefaultOrHelpMode(toggleBetweenMode)
                Else
                    GoTo pressedInError
                End If
            Case Else
pressedInError:
                ' Assume you are in a normal document and have pressed this in error
                msgBox("This macro allows the user to change the defualt values and can only be run from the " & _
                "Protected GoForm, such as W00-GoForm1.0.doc. You are not in a goForm.  No action has been performed.")
        End Select
        toggleDefaultOrHelpMode = NO_ERROR
        Exit Function

ErrorHandler:
        msgBox("Error in toggleDefaultOrHelpMode: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        toggleDefaultOrHelpMode = FUNCTION_ERROR
        Exit Function

    End Function

    Function enterDefaultOrHelpMode(ByVal desiredMode As String) As String
        ' enterDefaultOrHelpMode - Andrea Jan2008
        ' DESCRIPTION: Use for entering either change of starting go
        '   file default or help mode

        Dim status As String
        Dim newColor As Object
        Dim msg As String
        Dim newMode As String
        Dim formVersion As String

        On Error GoTo ErrorHandlerProtected
        formVersion = ActiveDocument.FormFields("formVersion").TextInput.Default
        formVersion = Left(formVersion, 1)

        If (formVersion = 1) Then
            ' do not enter mode.  Can no longer support change default in less than Version 2
            msg = "Editing Defaults or Help is no longer available in goForm version 1. " & _
                  "No action taken. "
            msgBox(msg, vbOKOnly + vbCritical)
            Exit Function
        End If

        Select Case desiredMode
            Case DEFAULT_MODE
                newColor = 11923927
                msg = "CHANGE DEFAULTS MODE entered. Press Ctrl+Shift+E again or close the window to exit this mode"
                newMode = DEFAULT_MODE
            Case HELP_MODE
                newColor = 6725352
                msg = "CHANGE HELP MODE entered. Press Ctrl+Shift+R again or close the window to exit this mode"
                newMode = HELP_MODE
            Case Else
                msgBox("error:  Invalid SW input to the function enterDefaultOrHelpMode")
                Exit Function
        End Select
        msgBox(msg)
        unprotectDocument()
        On Error GoTo ErrorHandlerUnprotected
        status = changeBackgroundColors(wdColorGray10, newColor)
        If (status <> NO_ERROR) Then GoTo ErrorHandlerUnprotected

        ActiveDocument.fields.Update()
        changeMode(newMode)
        protectDocument(wdAllowOnlyFormFields)
        On Error GoTo ErrorHandlerProtected
        If (desiredMode = HELP_MODE) Then
            'insert title so other than color, user knows what mode you are in.
            'change all on exit macros to be the 'bring up help screen'
            setAllFieldOnExitMacros("editFormFieldHelp")
        End If

        Exit Function
ErrorHandlerProtected:
        msgBox("Error in enterDefaultOrHelpMode protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        enterDefaultOrHelpMode = FUNCTION_ERROR
        Exit Function

ErrorHandlerUnprotected:
        msgBox("Error in enterDefaultOrHelpMode unprotected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        protectDocument(wdAllowOnlyFormFields)
        enterDefaultOrHelpMode = FUNCTION_ERROR
        Exit Function

    End Function

    Function exitDefaultOrHelpMode(ByVal desiredMode As String) As String
        ' exitDefaultOrHelpMode - Andrea Jan2008
        ' DESCRIPTION: Use for exiting either change of starting go
        '   file default or help mode

        Dim status As String
        Dim replaceColor As Object
        Dim msg As String
        Dim newMode As String

        On Error GoTo ErrorHandlerProtected
        newMode = START_MODE
        Select Case desiredMode
            Case DEFAULT_MODE
                replaceColor = 11923927
                msg = "EXITING CHANGE DEFAULTS MODE"
            Case HELP_MODE
                replaceColor = 6725352
                msg = "EXITING CHANGE HELP MODE."
            Case Else
                msgBox("error:  Invalid SW input to the function exitDefaultOrHelpMode")
                GoTo ErrorHandlerProtected
        End Select

        msgBox(msg)
        unprotectDocument()
        On Error GoTo ErrorHandlerUnprotected
        If (desiredMode = DEFAULT_MODE) Then
            'save entry of the field you are currently in
            setDefaultFormField()

        ElseIf (desiredMode = HELP_MODE) Then
            'change all exit macros back to 'normal' setDefaultFormFields
            setAllFieldOnExitMacros("setDefaultFormField")
        End If
        ActiveDocument.fields.Update()
        status = changeBackgroundColors(replaceColor, wdColorGray10)
        If (status <> NO_ERROR) Then GoTo ErrorHandlerUnprotected

        changeMode(newMode)
        protectDocument(wdAllowOnlyReading)
        exitDefaultOrHelpMode = NO_ERROR
        Exit Function

ErrorHandlerProtected:
        msgBox("Error in exitDefaultOrHelpMode protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        exitDefaultOrHelpMode = FUNCTION_ERROR
        Exit Function

ErrorHandlerUnprotected:
        msgBox("Error in exitDefaultOrHelpMode unprotected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)

        protectDocument(wdAllowOnlyFormFields)
        exitDefaultOrHelpMode = FUNCTION_ERROR
        Exit Function
    End Function

    Function changeBackgroundColors(ByVal originalColor As String, ByVal newColor As String)
        ' changeBackgroundColors - 2008Jan Andrea
        ' DESCRIPTION: This function changes the background colors of the goForm when editing
        '   defaults.  Gives the user the 'feel' that they are in a special mode of the goForms,
        '   and also lets them know when they have left it.

        Dim numMoved As Integer
        Dim count As Integer
        Dim backgroundColor As String

        On Error GoTo ErrorHandler
        count = 0
        ' go to start of document, and select the first line
        Selection.HomeKey Unit:=wdStory
        numMoved = Selection.MoveEnd(wdLine, 1)

        Do
            ' an error will occur when we reach the end of the file, so catch it in the error handler
            backgroundColor = Selection.ParagraphFormat.Shading.BackgroundPatternColor
            If (backgroundColor = originalColor) Then
                Selection.ParagraphFormat.Shading.BackgroundPatternColor = newColor
            End If

            ' collapse, move 1 right bringing us to next line, Highlight next line for next read
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.MoveRight(Unit:=wdCharacter, count:=1, Extend:=wdExtend)
            numMoved = Selection.MoveEnd(wdLine, 1)
            count = count + 1
        Loop Until count > 150

        Selection.HomeKey Unit:=wdStory
        changeBackgroundColors = NO_ERROR
        Exit Function
ErrorHandler:
        If (Err.Number = 91) Then
            ' We have finished changing all of the relevant lines in the file
            changeBackgroundColors = NO_ERROR
        Else
            ' All other errors.  Display error number & exit
            msgBox("Error in changeBackgroundColors: " & "number: " _
                    & Err.Number & "-message : " & Err.Description)
            changeBackgroundColors = FUNCTION_ERROR
        End If

    End Function

    Sub editFormFieldHelp()
        ' editFormFieldHelp - 2008Jan andrea
        ' DESCRIPTION:  This is used by enter&exiting changing form field HELP_MODE.
        '   This is an on exit macro that is assigned to all formfields when HELP_MODE
        '   is entered.  If the user has entered ** in the beginning of the form field,
        '   it asks the user to revise & save the contents of the help for that field.

        Dim currentBkm As String
        Dim currentHelpText As String
        Dim currentBkmResult As String
        Dim firstLetters As String
        Dim msgBoxTitle As String
        Dim returnedHelpText As String
        On Error GoTo ErrorHandler

        currentBkm = Selection.Bookmarks(1).name

        ' To change help on status bar, must set ownhelp to true
        ActiveDocument.FormFields(currentBkm).OwnHelp = True
        currentHelpText = ActiveDocument.FormFields(currentBkm).StatusText
        currentBkmResult = trimWhiteSpace(ActiveDocument.FormFields(currentBkm).result)
        firstLetters = Left(currentBkmResult, 1)

        ' only if the user has entered "**" in the first 2 field characters, do
        ' want to edit help.  Otherwise the user is just browsing what the help is
        If (firstLetters = "=") Then

            'bring up user input dialog
            msgBoxTitle = UCase$(currentBkm) & ": review & click ok to accept changes"
            returnedHelpText = askUserForInputBox("Help text", msgBoxTitle, currentHelpText)
            If (returnedHelpText = FUNCTION_ERROR) Then GoTo ErrorHandler

            If (returnedHelpText = USER_EXIT) Then
                Exit Sub
            Else
                'only update the help if the user did not press cancel in the msgBox
                ActiveDocument.FormFields(currentBkm).OwnHelp = True
                ActiveDocument.FormFields(currentBkm).StatusText = returnedHelpText
            End If
        End If

        Exit Sub
ErrorHandler:
        msgBox("Error in editFormFieldHelp: " & "number: " _
                & Err.Number & "-message : " & Err.Description)


        Exit Sub
    End Sub


End Module





