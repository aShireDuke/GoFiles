Module FormDocumentConstructionWills

Option Explicit On
    ' define global constant for error messages on input boxes, etc

    ' MODULE DESCRIPTION:  This module contains form routines that are specific to Wills.
    '   The following functions may use routines from the more general modules of
    '   FormDocumentConstruction and FormDocumentControls.
    ' *****************************************************************

    Sub willsSwitchHisToHers()
        ' willsSwitchHisToHers - Nov19_07 Andrea
        ' DESCRIPTION:  This function will be run in a unprotected saved will document.  When a husband and wife do
        '   up a will, only one go file is created, the will is created for the husband first, revised and finished.
        '   After which the will of the wife is created by changing the name and the gender references, etc.
        ' KEYBOARD SHORTCUT: Ctrl+shift+H (based on the H is his/hers)

        Dim numMoved As Integer
        Dim lineRead1, lineread2 As String
        Dim originalName, newName As String
        Dim correctName1, correctName2 As String
        Dim status As String
        Dim dialogStatus As String
        Dim fileName As String
        Dim position As Integer
        Dim doc As String

        ' Check if this macro is applicable -- are you in a nonProtected document of some sort?
        Dim protection As String
        protection = ActiveDocument.protectionType

        If (protection = wdAllowOnlyReading Or protection = wdAllowOnlyFormFields) Then
            'Have pressed this in error.  We only want to run this on an unlinked document
            msgBox("This document protection type does not support histo Hers Macro.  " & _
                "Exiting.  No action Performed")
            Exit Sub
        End If

        ' Suggest Original.  Move to End, up one section, and back to get
        ' to the starting character of the labelled u and uspouse
        Selection.EndKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToPrevious, count:=1, name:="")
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=1, name:="")

        ' move 2 lines past the histohersTitle, past label for u and collapse
        numMoved = Selection.MoveEnd(wdLine, 2)
        Selection.Collapse Direction:=wdCollapseEnd

        ' move to right, bringing you on to next line, which should be the value of u
        Selection.MoveRight(Unit:=wdCharacter, count:=1, Extend:=wdExtend)

        ' The originalName should be on this line. Move to the endofLine, back one so no formatting
        numMoved = Selection.MoveEnd(wdLine, 1)
        Selection.MoveLeft(Unit:=wdCharacter, count:=1, Extend:=wdExtend)
        lineRead1 = Trim$(Selection.Text)

        ' move to end of this line (incl formatting), move to end of next line (label for uSpouse)
        ' Collapse, then move to end of next line (value of uSpouse) & back one char to not include formatting
        numMoved = Selection.MoveEnd(wdLine, 2)
        Selection.Collapse Direction:=wdCollapseEnd
        numMoved = Selection.MoveEnd(wdLine, 1)
        Selection.MoveLeft(Unit:=wdCharacter, count:=1, Extend:=wdExtend)
        lineread2 = Trim$(Selection.Text)

        ' Go home so can see the beggining of the will to check visually against names there
        Selection.HomeKey Unit:=wdStory
        correctName1 = askUserForInputBox("Is this the correct Name?", _
                                          "Is this the name of the testator in this document, the original will? ", lineRead1)
        If (correctName1 = USER_EXIT) Then
            msgBox("Error.  His to Hers Macro exited by user or name input is blank.  No replacements made")
            Exit Sub
        ElseIf (correctName1 = FUNCTION_ERROR) Then
            GoTo ErrorHandler
        End If
        originalName = Trim$(correctName1)

        correctName2 = askUserForInputBox("Is this the correct Name?", _
                                          "Is this the name of the spouse in this document, the original will", lineread2)
        If (correctName2 = USER_EXIT) Then
            msgBox("Error.  His to Hers Macro exited by user or name input is blank.  No replacements made")
            Exit Sub
        ElseIf (correctName2 = FUNCTION_ERROR) Then
            GoTo ErrorHandler
        End If
        newName = Trim$(correctName2)

        ' Now we have the originalName and the NewName - Save as then Go replace the values
        ' we do not want the last .doc part of the active document name.
        fileName = ActiveDocument.name
        doc = ".doc"
        position = InStr(1, fileName, doc)
        fileName = Left$(fileName, position - 1)
        dialogStatus = saveDialog(fileName & "-spouse", ActiveDocument.path)
        If (dialogStatus = USER_EXIT) Then
            msgBox("His to Hers Macro exited by user.  No replacements made")
            Exit Sub
        ElseIf (dialogStatus = FUNCTION_ERROR) Then
            msgBox("Error in dialog in willsSwitchHisToHers.  No replacements made. Exiting..")
            Exit Sub
        End If

        ' Find and replace all instances of value of originalName with the
        ' string, "originalName" & value of newName with string "newName"
        Selection.HomeKey Unit:=wdStory
        status = findAndReplaceCustom(originalName, "originalName")
        status = findAndReplaceCustom(newName, "newName")

        ' Find and replace "originalName" with the value of newName
        ' Find and replace "newName" with the value of "originalName"
        status = findAndReplaceCustom("originalName", newName)
        status = findAndReplaceCustom("newName", originalName)
        Exit Sub

ErrorHandler:
        msgBox("Error in willsSwitchHisToHers: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        Exit Sub

    End Sub

    Function askUserForInputBox(ByVal title As String, ByVal nameQuestion As String, _
                                ByVal name As String) As String
        ' askUserIfCorrectName - 2007Dec - Andrea
        ' DESCRIPTION: This function is used by willsSwitchHisToHers.  Brings up a dialog box to
        '   with a suggestion of a name, and asks the user to modify or confirm.

        Dim checkedName As String
        On Error GoTo ErrorHandler

        checkedName = InputBox(Prompt:=nameQuestion, title:=title, Default:=name)
        If (checkedName = "") Then
            askUserForInputBox = USER_EXIT
        Else
            askUserForInputBox = checkedName
        End If

        Exit Function
ErrorHandler:
        msgBox("Error in askUserForInputBox: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        askUserForInputBox = FUNCTION_ERROR
    End Function

    Function findAndReplaceCustom(ByVal replaceVar As String, ByVal varName As String) As String
        ' findAndReplaceCustom - 2007Dec - Andrea
        ' DESCRIPTION: Uses the word find and replace with the settings specified below.

        On Error GoTo ErrorHandler
        With Selection.Find
            .Text = replaceVar
            .Replacement.Text = varName
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        findAndReplaceCustom = "ok"
        findAndReplaceCustom = NO_ERROR
        Exit Function

ErrorHandler:
        msgBox("Error in findAndReplaceCustom: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        findAndReplaceCustom = FUNCTION_ERROR
    End Function

    Function insertHisHersWillInformation(ByVal titleString) As String
        ' insertHisHersWillInformation - Nov30_07 Andrea
        ' DESCRIPTION:  This function inserts the main name (u) and the spouse's name
        '   at the bottom of any merged will file for His to hers conversion
        On Error GoTo ErrorHandler
        Selection.TypeText(titleString)
        Selection.TypeParagraph()
        Selection.TypeText("u: ")
        Selection.TypeParagraph()
        Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="Ref u \* charformat")
        Selection.TypeParagraph()
        Selection.TypeText("uSpouse: ")
        Selection.TypeParagraph()
        Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
        Text:="Ref uSpouse \* charformat")
        insertHisHersWillInformation = NO_ERROR
        Exit Function
ErrorHandler:
        msgBox("Error in insertHisHersWillInformation: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        insertHisHersWillInformation = FUNCTION_ERROR
    End Function

End Module