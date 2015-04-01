Module General

Option Explicit On
    ' MODULE: General.  This module contains the general day to day use macros that
    '   a user would need in word.  This includes findStar, letterheadChange, PCP,
    '   etc.  Eventually if this module grows large enough consider breaking into more
    '   modules such as formatting, printing, etc
    '***********************************************************

    Sub findStar()
        ' FindStar - 04Jan1,mod 04May29 - Grant Dukeshire
        ' DESCRIPTION: The star is used as a wild card entry for editing, document production, etc.
        '   This macro allows the user to jump from star(*) to star(*) within a regular or protected
        '   document.
        ' KEY SHORTCUT: F12 & Alt&A. Reason as used Most frequently (& F10,F11 harmless).  Then if in a GoFile,
        '   the sequence is F9-F10-F11-F12-CF12 to merge work

        Dim strProtectionType As String

        On Error GoTo ErrorHandler
        strProtectionType = ActiveDocument.protectionType

        Select Case strProtectionType
            Case wdAllowOnlyFormFields
                findStarWithinFormField()
            Case wdNoProtection
                findStarNoProtection()
            Case Else
                MsgBox("Action not allowed.  You can not find a star in a protected read only document.  See Macro " _
                & "findStar in General Module for more details")
        End Select
        Exit Sub
ErrorHandler:
    msgBox ("Error number: " & Error.Number & " description: " & Error.Description)

    End Sub
    Function findStarNoProtection()
        ' FindStarNoProtection - James, modified Andrea Nov19_07
        ' DESCRIPTION: This function is called from FindStar to find a star within a non protected
        '       document, such as old Gofiles (ie W00-Go1.doc) or unlinked precedent files.

        On Error GoTo ErrorHandler
        If Selection.Words.count > 1 Then      ' Collapses if any selection first made
            Selection.MoveLeft(Unit:=wdCharacter, count:=1)
        End If
        With Selection.Find
            .Text = "*"
            .Wrap = wdFindContinue
            .Forward = True
            .Execute()
        End With
        Exit Function
ErrorHandler:
    msgBox ("Error number: " & Error.Number & " description: " & Error.Description)

    End Function

    Function findStarWithinFormField()
        ' FindStarWithinFormField - James, some mod Andrea Nov19_07
        ' DESCRIPTION: This function is called from FindStar to find a star within a
        '       protected document, intended to be a document containing formfields.

        Dim length As String
        Dim Initleng As String
        Dim myRange
        myRange = Selection.Range
        Dim RawString, StrClean1 As String

        ' update the form field value you are currently in
        setDefaultFormField()

        ' find the length of range you have selected.  Initialize the starting length
        Initleng = myRange.End - myRange.Start
        length = Initleng

        ' InStr function returns the position of the first occurrence of a string in another string.
        ' The syntax for the InStr function is:
        ' InStr( [start], string_being_searched, string2, [compare] )

        ' search for '*' in the string read
        RawString = Selection
        StrClean1 = InStr(1, RawString, "*")

        ' If a cset, which is used to search for a specified thing, is executed at the last
        ' position of a formfield it will crash and freeze the program.
        ' That is why there is always a search ahead before it looks further

        If (length = 1 Or length = 0) Then
            If (length = 1) Then
                ' this line makes the cursor go to the next field if at end of field
                Selection.MoveRight(Unit:=wdWord, count:=1, Extend:=wdMove)
            End If

            Selection.MoveRight(Unit:=wdWord, count:=100, Extend:=wdExtend)
            RawString = Selection
            StrClean1 = InStr(1, RawString, "*")
            If StrClean1 = 0 Then
                Selection.Collapse Direction:=wdCollapseEnd
                GoTo EndofField
            End If
            Selection.Collapse()
            Selection.MoveUntil(Cset:="*", count:=wdForward)
            Selection.MoveRight(Unit:=wdCharacter, count:=1, Extend:=wdExtend)
            Exit Function

        Else
            'For all other lengths
            RawString = Selection
            StrClean1 = InStr(1, RawString, "*")
            If StrClean1 = 0 Then GoTo EndofField
            Selection.Collapse()
            Selection.MoveUntil(Cset:="*", count:=wdForward)
            Selection.MoveRight(Unit:=wdCharacter, count:=1, Extend:=wdExtend)

            myRange = Selection.Range

            length = myRange.End - myRange.Start
            Dim Comboleng As String
            Comboleng = Initleng - length

            If Comboleng = 0 Then
                If Initleng = 1 Then Exit Function
                Selection.Collapse Direction:=wdCollapseEnd
                'findStarWithinFormField
                GoTo EndofField
            End If
        End If

        Exit Function
        ' So now that we've gotten to here, go to next field
EndofField:
        unprotectDocument()
        On Error GoTo EndOfDocumentErrorHandler
        setDefaultFormField()
        Selection.NextField.Select()
        GoTo ResumeProtection

ResumeProtection:
        Selection.Collapse Direction:=wdCollapseStart
        protectDocument(wdAllowOnlyFormFields)
        Exit Function

EndOfDocumentErrorHandler:
        If (Err.Number = 91) Then
            ' End of file Error. Go to beginning and search there
            setDefaultFormField()
            Selection.HomeKey Unit:=wdStory
            Selection.NextField.Select()
            Selection.Collapse Direction:=wdCollapseStart
            GoTo ResumeProtection
        Else
            ' All other errors
            MsgBox("Custom Error:  Error in function FindStarWithinFormField")
            GoTo ResumeProtection
        End If

    End Function

    Sub letterheadChange()
        ' LetterheadChange - 04Dec20,mod 06Aug17 - Grant Dukeshire
        ' DESCRIPTION: This function is used when we have a new lawyer or student at law.  Replaces
        '   all letterheads of (KT,B&S & KT,S-a-L) to Erin Barvir, S-a-L.  Useful during the
        '   transition period when are still opening documents that had the lawyer with the
        '   previous title.
        ' KEY SHORTCUT: Alt-F5 .Reason: as F5 is just ToolsCalc (harmless)
        With Selection.Find
            .Text = "Kathryn Tweedie, Student-at-Law^tFax"
            .Replacement.Text = "Erin Bavir, Barrister & Solicitor^tFax"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .Text = "Kathryn Tweedie, Barrister & Solicitor^tFax"
            .Replacement.Text = "Erin Barvir, Barrister & Solicitor^tFax"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .Text = "Arlene K. Blake, Student-at-Law^tFax"
            .Replacement.Text = "Erin Barvir, Barrister & Solicitor^tFax"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .Text = "Erin Barvir, Student-at-Law^tFax"
            .Replacement.Text = "Erin Barvir, Barrister & Solicitor^tFax"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .Text = "Erin Barvir, Barrister & Solicitor (on leave - inactive)^tFax"
            .Replacement.Text = "Erin Barvir, Barrister & Solicitor^tFax"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
        With Selection.Find
            .Text = "Erin Barvir, Barrister & Solicitor^tFax:  403-286-7644"
            .Replacement.Text = "Erin Barvir, Barrister & Solicitor (on leave - inactive)^tFax:  403-286-7644" & Chr$(13) & "Fanny Deng, Student - at - Law"
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft

    End Sub

    Sub printCurrentPage()
        ' PrintCurrentPage Macro - 04July31 - Grant Dukeshire
        ' DESCRIPTION: Use to print the current page without going through the print menus with the mouse
        ' KEYSHORTCUT: None.  Reason shortcut key assigned, in toolbar
        On Error GoTo ErrorHandlerPCP
        Application.PrintOut(fileName:="", Range:=wdPrintCurrentPage, Item:= _
            wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
            ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
            False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
            PrintZoomPaperHeight:=0)
        Exit Sub
ErrorHandlerPCP:
        Dim msg As String
        msg = "An error occured. " & "Error number: " & Err.Number & _
               "Error message: " & Err.Description
        MsgBox(msg, vbOKOnly + vbCritical)

    End Sub

    Sub toggleCodes()
        ' ToggleCodes - 05June30 - Grant Dukeshire
        ' DESCRIPTION: Toggles the view between showing the field codes (go file entry) and showing
        '   the results of the update.  This macro Enables getting back to & from GoFile or Results views
        ' KEYSHORTCUT: F10 key.  Reason as often used after F9 updating
        ' Will not work if the document is protected
        ActiveDocument.fields.ToggleShowCodes()
    End Sub

    Sub showSpellErrors()
        ' 05July - James Dukeshire
        ' DESCRIPTION: This function can be run in many old documents that have proofing errors.
        ' KEY SHORTCUT: F7 as this is the normal spell checking key
        ActiveDocument.Range.NoProofing = False
    End Sub
    Sub calcPaste()
        '   05May20 by James Dukeshire
        '   DESCRIPTION: In word, you can highlight numbers and press "ctrl=" to call the built in
        '       Word function 'ToolsCalculate' to calculate the highlighted arithmetic.  The result
        '       of the arithmetic is pasted into the clipboard, and the user presses Ctr+v to insert.
        '       THIS function was created to fix a formatting problem with word 2003, to ensure the
        '       number calculated by 'ToolsCalculate' has double precision after the decimal and a
        '       comma to separate the thousands such as the number 33,333.00.  To use this macro
        '       you use 'Ctrl-=' and then 'Ctrl-[' to paste
        '   KEY SHORTCUTS: "Ctrl-[", "Ctrl-]", "Alt-[", "Alt-[".  Reason as are just under "=" on keyboard
        '   Key shortcut to Word Command ToolsCalculate = "Ctrl=", "Alt=" & on toolbar as 'Calc'
        '***********************************************************

        Dim MyData, ClipData
        MyData = New DataObject
        ClipData = New DataObject
        Dim RawNum As Double
        Dim StrRawNum, NewNum As String

        On Error GoTo ErrorHandler
        ClipData.GetFromClipboard()
        RawNum = ClipData.GetText
        ' Convert the rawNum to a string
        StrRawNum = CStr(RawNum)
        NewNum = FormatNumber(StrRawNum, 2)
        StatusBar = NewNum

        ' Puts the reformatted ToolsCalculate result and puts into clipboard, so it
        ' can be pasted later
        MyData.SetText(NewNum)
        MyData.PutInClipboard()
        Selection.Paste()
        Exit Sub

ErrorHandler:
        Dim msg As String
        msg = "An error occured! " & "Error number: " & Err.Number & "Error message: " & Err.Description
        MsgBox(msg, vbOKOnly + vbCritical)

    End Sub


End Module





























