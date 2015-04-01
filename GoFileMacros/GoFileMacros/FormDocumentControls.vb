Module FormDocumentControls

Option Explicit On
    ' MODULE DESCRIPTION:  This module contains functions and subroutines that are neccesary for
    '   navigating and saving the go form.  Applicable actions include what happens when a general
    '   button is pushed, and what happens when the user exits a form field.  Functions placed in
    '   this module should apply to all areas of law, and be the core field control of the Go
    '   form functions
    '***********************************************************


    Sub ProgrammerSetAllFieldsToDefaultSave()
        ' ProgrammerSetAllFieldsToDefaultSave - Andrea Jan 2008
        ' DESCRIPTION:  Used by the programmer to do a global set of the field exit macros to
        '   setDefaultFormField.  Used if the form has gone to a 'odd' state, or if creating
        '   a new form.
        setAllFieldOnExitMacros("setDefaultFormField")

    End Sub

    Function setAllFieldOnExitMacros(ByVal exitMacro As String) As String
        ' setAllFieldsOnExitMacros - Dec17, 2007 Andrea
        ' DESCRIPTION:  This function is used by the programmer when setting up a new
        '   goForm from scratch.  It must be run manually when all protection is removed
        '   from the document.  This function goes through all of the fields in the
        '   document and sets each exitMacro to the same function, neccesary for the
        '   data to be able to be stored in the GoForm.

        Dim J As Integer
        Dim bkm As String
        On Error GoTo ErrorHandler

        For J = 1 To ActiveDocument.Bookmarks.count
            bkm = ActiveDocument.Bookmarks(J).name
            ActiveDocument.FormFields(bkm).exitMacro = exitMacro
        Next J
        setAllFieldOnExitMacros = NO_ERROR
        Exit Function
ErrorHandler:
        MsgBox("Error in setAllFieldOnExitMacros : " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        setAllFieldOnExitMacros = FUNCTION_ERROR
        Exit Function

    End Function

    Sub setDefaultFormField()
        ' setDefaultFormField - Dec 2007 andrea
        ' DESCRIPTION: This function is the generic exit function for ALL formfields in
        '   the goFile.  It reads the name of the bookmark that the user is currently
        '   in or has selected, updates the bookmark in the document to the value inputted,
        '   and then writes the newly inputted information into the default value for the
        '   field (so when open the document next time, the default (now the specific client
        '   information) is displayed/saved in the formfields.

        Dim currentBkm As String
        Dim trimmedString As String

        On Error GoTo ErrorHandler
        currentBkm = Selection.Bookmarks(1).name
        trimmedString = trimWhiteSpace(ActiveDocument.Bookmarks(currentBkm).Range)
        ActiveDocument.FormFields(currentBkm).TextInput.Default = trimmedString
        ActiveDocument.Bookmarks(currentBkm).Range.fields(1).Update()
        Exit Sub

ErrorHandler:
        Dim msg As String
        msg = "Error in setDefaultFormField: " & "number: " _
                & Err.Number & "-message : " & Err.Description
    End Sub

    Sub calculateDescription()
        ' calculateDescription - 2008 January - Andrea
        ' DESCRIPTION:  This sub is used as an on entry macro for uDesc, xDesc,
        '   and re2.  It can also be used to set xAd1, xAd2, uAd1, uAd2 to ad1,
        '   and ad2, but for now we are going back to the old blank method.

        Dim msgPrompt As String
        Dim msgTitle As String
        Dim msgResult As Integer
        Dim currentBkm As String
        Dim currentValue As String
        Dim ad1, ad2, ad3 As String
        Dim result As String

        On Error GoTo ErrorHandler
        currentBkm = Selection.Bookmarks(1).name
        currentValue = trimWhiteSpace(ActiveDocument.Bookmarks(currentBkm).Range)

        ' We only want to ask the user if they want the computer to generate if the
        ' current value of the bookmark is blank.  This way only asks the first time,
        ' or if the user 'forces' the SW to ask by deleting it all
        If (currentValue <> "     ") Then
            Exit Sub
        End If

        ' Case statment for which Message box to bring up to the user
        Select Case currentBkm
            Case "uDesc", "xDesc", "re2"
                msgPrompt = "Do you want to Auto Fill?" & _
                    vbCr & vbCr & _
                    "Yes: Have program generate from goFile" & vbCr & _
                    "No: User will input.  No action taken"
                msgTitle = "Generate a Description"
                msgResult = customMsgBox(msgPrompt, msgTitle)

            Case "xAd1", "uAd1"
                msgPrompt = "Do you want to use Ad1 & Ad2 for the title address?" & _
                    vbCr & vbCr & _
                    "Yes: Use Ad1 & Ad2 inputted in this goFile" & vbCr & _
                    "No: No action taken.  User will input"
                msgTitle = "Address to be shown on title"
                msgResult = customMsgBox(msgPrompt, msgTitle)

            Case Else
                MsgBox("Custom Error: This macro should not be connected to this field.  See " & _
                        "current field properties under entry macro")
                Exit Sub
        End Select

        'case statment for what to do with the custom message box that the user answered for
        Select Case msgResult
            Case vbYes
                Select Case currentBkm
                    Case "uDesc"
                        result = getDescriptionFromBookmarks("u1", "u2", "uAd1", "uAd2")
                        If (result = FUNCTION_ERROR) Then GoTo ErrorHandler
                        ActiveDocument.FormFields(currentBkm).result = result
                    Case "xDesc"
                        result = getDescriptionFromBookmarks("x1", "x2", "xAd1", "xAd2")
                        If (result = FUNCTION_ERROR) Then GoTo ErrorHandler
                        ActiveDocument.FormFields(currentBkm).result = result
                    Case "re2"
                        result = getReLineFromBookmarks("ad1", "ad2", "legal1")
                        If (result = FUNCTION_ERROR) Then GoTo ErrorHandler
                        ActiveDocument.FormFields(currentBkm).result = result
                    Case "uAd1"
                        ad1 = trimWhiteSpace(ActiveDocument.Bookmarks("ad1").Range)
                        ad2 = trimWhiteSpace(ActiveDocument.Bookmarks("ad2").Range)
                        ad3 = trimWhiteSpace(ActiveDocument.Bookmarks("ad3").Range)
                        ActiveDocument.FormFields(currentBkm).result = ad1
                        ActiveDocument.FormFields("uAd2").result = ad2 & ", " & ad3
                    Case "xAd1"
                        ad1 = trimWhiteSpace(ActiveDocument.Bookmarks("ad1").Range)
                        ad2 = trimWhiteSpace(ActiveDocument.Bookmarks("ad2").Range)
                        ad3 = trimWhiteSpace(ActiveDocument.Bookmarks("ad3").Range)
                        ActiveDocument.FormFields(currentBkm).result = ad1
                        ActiveDocument.FormFields("xAd2").result = ad2 & ", " & ad3
                End Select
            Case vbNo
                Exit Sub
        End Select
        Exit Sub
ErrorHandler:
        Dim msg As String
        msg = "Error in calculateDescription: " & "number: " _
                & Err.Number & "-message : " & Err.Description

    End Sub
    Function customMsgBox(ByVal msgPrompt As String, ByVal msgTitle As String) As Integer
        ' customMsgBox - Andrea 2008Jan
        ' DESCRIPTION:  This function is primarily used by calculateDescription.  It can be used
        '   when the SW needs to ask the user a yes or no question.  The return value is
        '   returned from custom msgbox and is either the string "vbYes" or "vbNo"

        Dim msgButtons As Integer

        msgButtons = vbYesNo + vbQuestion + vbDefaultButton2
        customMsgBox = MsgBox(msgPrompt, msgButtons, msgTitle)
        Exit Function

ErrorHandler:
        MsgBox("Error in customMsgBox: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        customMsgBox = FUNCTION_ERROR
    End Function


    Function getReLineFromBookmarks(ByVal Ad1Bkm As String, ByVal Ad2Bkm As String, _
                                    ByVal legalBkm As String) As String
        ' getReLineFromBookmarks - Andrea 2008Jan
        ' DESCRIPTION:  This function is used by t
        Dim ad1 As String
        Dim ad2 As String
        Dim legal As String
        Dim temp As String
        On Error GoTo ErrorHandler

        ad1 = trimWhiteSpace(ActiveDocument.Bookmarks(Ad1Bkm).Range)
        ad2 = trimWhiteSpace(ActiveDocument.Bookmarks(Ad2Bkm).Range)
        legal = trimWhiteSpace(ActiveDocument.Bookmarks(legalBkm).Range)

        temp = ad1 & ", " & ad2 & " - " & legal
        getReLineFromBookmarks = temp
        Exit Function

ErrorHandler:
        MsgBox("Error in getReLineFromBookmarks: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        getReLineFromBookmarks = FUNCTION_ERROR
    End Function
    Function getDescriptionFromBookmarks(ByVal p1Bkm As String, ByVal p2Bkm As String, _
                                         ByVal Ad1Bkm As String, ByVal Ad2Bkm As String) As String
        ' getDescriptionFromBookmarks - Jan 08 Andrea
        ' Description: Takes in names of bookmarks neccesary for constructing a description.  This is
        '   used by both uDesc and xDesc.  Function written so the calling function calculateDescription
        '   does not have to repeat the same code for essentially the same thing

        'used p1, p2 for person1, person2.  So same if is x or u
        Dim p1 As String
        Dim p2 As String
        Dim ad1 As String
        Dim ad2 As String
        Dim ad3 As String
        Dim pAd1 As String
        Dim pAd2 As String
        Dim temp As String

        p1 = trimWhiteSpace(ActiveDocument.Bookmarks(p1Bkm).Range)
        p2 = trimWhiteSpace(ActiveDocument.Bookmarks(p2Bkm).Range)
        pAd1 = trimWhiteSpace(ActiveDocument.Bookmarks(Ad1Bkm).Range)
        pAd2 = trimWhiteSpace(ActiveDocument.Bookmarks(Ad2Bkm).Range)
        ad1 = trimWhiteSpace(ActiveDocument.Bookmarks("ad1").Range)
        ad2 = trimWhiteSpace(ActiveDocument.Bookmarks("ad2").Range)
        ad3 = trimWhiteSpace(ActiveDocument.Bookmarks("ad3").Range)

        ' If the user has left x, uAd1 blank, they intend to use ad1, ad2, as set
        ' at time of merging by the calc function/file.  So use it here.
        ' This is made complex/odd because uAd2 has the postal code, whereas ad2, ad3
        ' are separate in the code.
        If (pAd1 = "     " Or pAd1 = "") Then
            pAd1 = ad1
            pAd2 = ad2 & "  " & ad3
        End If

        If (p2 = "" Or p2 = "     ") Then
            temp = p1 & " of " & pAd1 & ", " & pAd2
        Else
            temp = p1 & " and " & p2 & " both of " & pAd1 & ", " & pAd2 & " As Joint Tenants"
        End If

        getDescriptionFromBookmarks = temp
        Exit Function

ErrorHandler:
        MsgBox("Error in getDescriptionFromBookmarks: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        getDescriptionFromBookmarks = FUNCTION_ERROR
    End Function

    Function updateAllBookmarks()
        ' updateAllBookmarks - Jan2008 andrea
        ' DESCRIPTION: This function looks through all bookmarks in the file, and updates them

        Dim J As Integer
        Dim bkm As String
        Dim trimmedString As String
        On Error GoTo ErrorHandler

        'Set value of formfield you are currently in
        setDefaultFormField()

        For J = 1 To ActiveDocument.Bookmarks.count
            bkm = ActiveDocument.Bookmarks(J).name
            If (ActiveDocument.Bookmarks.Exists("bkm") = False) Then GoTo endOfFor
            trimmedString = trimWhiteSpace(ActiveDocument.Bookmarks(bkm).Range)
            ActiveDocument.FormFields(bkm).TextInput.Default = trimmedString
endOfFor:
        Next J
        updateAllBookmarks = NO_ERROR
        Exit Function

ErrorHandler:
        MsgBox("Error in updateAllBookmarks: " & "number: " _
                & Err.Number & "-message : " & Err.Description & _
                "The bookmark: " & bkm)
        updateAllBookmarks = FUNCTION_ERROR
    End Function


End Module








