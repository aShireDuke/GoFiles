Module FormDocumentConstruction

Option Explicit On
    ' MODULE DESCRIPTION:  This module contains the macros that are neccesary to construct a document
    '   from a form Go file.  This includes the actions of finding the precedents requested, updating
    '   fields, opening a new document with the constructed/merged file within it, asking the user
    '   to save to a predefined drive, etc.  Functions placed in this module should apply to all
    '   areas of law, and be the core merging/driving functions.
    '***********************************************************


    Sub FormDocumentConstructionAction()
        'FormDocumentConstructionAction - Nov 22, 2007 Andrea
        ' DESCRIPTION:  This function combines all actions that a user of the new FORM document
        '   construction package would use into one function.  These functions include SaveAsGo,
        '   UpdateAll, unlinkAllSaveAs, and no action if pressed in error.  This function decides
        '   what the user would like to do based on document properties - what is in the active Document
        '   protection.  As protection has never been used before in this office, using it to
        '   be able to identify whether a file is a goFile precedent, or a client Go file or a regular
        '   document is a very powerful tool to use for decision making and error handling.

        Dim protection As String
        Dim currentMode As String
        Dim status As String

        On Error GoTo ErrorHandler
        protection = getProtection()

        If (protection <> wdAllowOnlyReading And protection <> wdAllowOnlyFormFields) Then
            ' then know we are either in an old style document, or we are in a regular document.
            Application.Run MacroName:="UpdateAll"
        End If

        currentMode = getMode()

        ' Case using the current mode as decision.  Each case ensures that the protection
        '   aligns properly with the current mode, otherwise throws an error and exits.
        Select Case currentMode
            Case START_MODE
                ' Start mode is the protected, read only goForm that the user does a
                ' "special save as" to, to make the form save with the client name &
                '  editable with the information.
                If (protection <> wdAllowOnlyReading) Then GoTo ErrorHandler
                status = saveAsGoForm()

            Case HELP_MODE, DEFAULT_MODE
                If (protection <> wdAllowOnlyFormFields) Then GoTo ErrorHandler
                ' Help mode, and default mode are set up so that the user can change the
                ' defaults and help screens of the go file.  This mode can only be entered from
                ' the protected read only go file, as it saves its changes here.  If a user
                ' tries to press F3 here, they have pressed it in error as it would make no sense
                ' to merge a document from this point, you need to be in start mode to do that.
                MsgBox("Custom Error: F3 has been pressed in error.  This file is locked " & _
                        "for editing of defaults or help screens.  Exiting.  No action performed.")

            Case CLIENT_MODE
                ' Client mode is the document that holds all of the client information, for
                ' example, Smith-Go1.doc.  If you are in this file and press F3, you want to merge
                ' the files selected in the xxx Insert file choice section.
                updateAllGoForm()
            Case Else

PressInError:
                ' This is the else mode case, or the generic this button has been pressed in error.
                MsgBox("Custom Error:  F3 has been pressed in error.  This file is not in " & _
                        "the correct mode or protection combination.  No action taken.  " & _
                        "Exiting...")
        End Select
        Exit Sub

ErrorHandler:
        Dim msg As String
        Dim errNum As Integer
        errNum = Err.Number

        Select Case errNum
            Case 5941
                ' Know this error is caused by member of collection does not exist.  Most likely
                ' in this function when trying to access formfield "goMode", in a file that does
                ' not have this bookmark
                GoTo PressInError
                Exit Sub
            Case Else
                MsgBox("Error in FormDocumentConstructionAction: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        End Select

        Exit Sub
    End Sub

    Function saveAsGoForm() As String
        ' SaveAsGoForm - Nov19_07, Andrea
        ' DESCRIPTION:  This function saves the go form in the proper style depending on the
        '   variable formType as read from the form.
        ' KEY SHORTCUT: F3

        ' The following lines get the type, and version of the Go form.
        ' The formType (ie Will, Estate, etc) is read from the default text of the two formfields
        ' contained in the title of the GoForm itself.  When updating the version the default text
        ' of the formField containing the version number should be changed.

        Dim formType, formVersion, goFileLetter, goFileName As String
        Dim clientName As String
        Dim dialogStatus As String

        On Error GoTo ErrorHandler

        formType = ActiveDocument.FormFields("formType").TextInput.Default
        formVersion = ActiveDocument.FormFields("formVersion").TextInput.Default
        goFileLetter = Left(formType, 1)
        goFileName = goFileLetter & "Go" & formVersion

        clientName = askLastNameAndTrim()
        If clientName = USER_EXIT Then GoTo Error_User_exit_before_unprotect
        unprotectDocument()
        dialogStatus = saveDialog(clientName & "-" & goFileName, ActiveDocument.path)

        If (dialogStatus = USER_EXIT) Then GoTo Error_User_exit_after_unprotect

        dialogStatus = insertGoHistory("--", "Client Go File created", " ", " ", " ", " ")
        'update goMode field with new mode - must do in unprotected mode
        changeMode(CLIENT_MODE)
        protectDocument(wdAllowOnlyFormFields)

        Exit Function

ErrorHandler:
        Dim msg As String
        Dim errNum As Integer
        errNum = Err.Number

        Select Case errNum
            Case 0
Error_User_exit_after_unprotect:
                protectDocument(wdAllowOnlyReading)
Error_User_exit_before_unprotect:
                ' SaveAsForm Aborted midway
                MsgBox("Custom Error: SaveAs GoForm exited By user or input left blank. " & _
                        "File not saved.  Exiting")

                saveAsGoForm = USER_EXIT
            Case Else
                ' all other errors.
                MsgBox("Error in saveAsGoForm: " & "number: " _
                    & Err.Number & "-message : " & Err.Description)
                saveAsGoForm = FUNCTION_ERROR
        End Select
    End Function

    Public Function protectDocument(ByVal protectionType As String)
        ' protectDocument - 2007Dec - Andrea
        ' DESCRIPTION:  Protects the current document to the protectionType specified
        '   in the input field.  Important to set NoReset = true, to ensure that when reprotect
        '   the formfield, it does not delete the current value of the formfield you are in.
        On Error GoTo ErrorHandler

        ActiveDocument.protect(Password:="LawOffice", Type:=protectionType, NoReset:=True)

        Exit Function
ErrorHandler:
        MsgBox("Error in protectDocument: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        Exit Function
    End Function

    Public Function unprotectDocument()
        ' unprotectDocument - 2007Dec - Andrea
        ' DESCRIPTION:  Unprotects the current document.
        On Error GoTo ErrorHandler

        ActiveDocument.unprotect Password:="LawOffice"
        Exit Function
ErrorHandler:
        MsgBox("Error in unprotectDocument: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        Exit Function
    End Function

    Function askLastNameAndTrim() As String
        ' askLastNameAndTrim - Dec 2007 - Andrea
        ' DESCRIPTION: This function is called when the user is initially saving a file.  The
        '   function brings up a dialog box to ask for the clientName that we would like to save
        '   the doc to, trims blanks, and returns the name to the calling function
        Dim clientName, lastNameForm As String
        Dim StrClean1, StrClean2, StrClean3 As Object

        On Error GoTo ErrorHandler

        clientName = InputBox("Enter Client LastNameID to save", , "LastNameID")
        clientName = Trim$(clientName)

        ' Check client name for invalid filename characters of \,/,-, or no input
        StrClean1 = InStr(1, clientName, "/")
        StrClean2 = InStr(1, clientName, "\")
        StrClean3 = InStr(1, clientName, "-")
        If StrClean1 <> 0 Or StrClean2 <> 0 Or StrClean3 <> 0 Then
            MsgBox("Error = LastNameID cannot have a '/' or '\'or '-'.  File not saved")
            Exit Function
        End If
        If clientName = "" Or IsNull(clientName) Then
            askLastNameAndTrim = USER_EXIT
        Else
            askLastNameAndTrim = clientName
        End If
        Exit Function

ErrorHandler:
        MsgBox("Error in askLastNameAndTrim: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        askLastNameAndTrim = FUNCTION_ERROR
        Exit Function
    End Function

    Function saveDialog(ByVal suggestedName As String, ByVal savePath As String) As String
        ' saveDialog: andrea Nov20_07
        ' DESCRIPTION:  used by the macro saveAsGoForm to bring up the save dialog and ask
        '   for the appropriate client name, and save in the proper file name format

        ' Following Opens the normal SaveAs window and suggests a default input name
        ' being the default path where the InitialFileName GoFile resides that was opened
        ' If a BOGUS path is given it will go back to the last path that (default path)
        ' the text after the last "\" in the path is what is used as the default name
        Dim dlgSaveAs As FileDialog
        On Error GoTo ErrorHandler
        Dim action As Integer

        dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
        dlgSaveAs.InitialFileName = savePath & "\" & suggestedName
        dlgSaveAs.InitialView = msoFileDialogViewDetails
        dlgSaveAs.title = "SaveAs Dialog for Functions Macros saveAsGoForm and UpdateAllForm"
        dlgSaveAs.ButtonName = "Save"
        action = dlgSaveAs.Show
        ' if user presses save, action = 1, if presses cancel, action = 0
        If (action = -1) Then
            dlgSaveAs.Execute()
            saveDialog = NO_ERROR
        Else
            'The function goes here if there are problems withthe save dialog such as a user pressing
            'cancel save.  Both saveAsGoForm and updateAllGoForm use this dialog. Therefore must exit
            'with a error value of 0
            saveDialog = USER_EXIT ' was0
        End If
        Exit Function

ErrorHandler:
        MsgBox("Error in exitDefaultOrHelpMode protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        saveDialog = FUNCTION_ERROR ' was nonexistant
        Exit Function
    End Function

    Function updateAllGoForm()
        ' updateaAllGoForm - Nov 20_07 Andrea
        ' DESCRIPTION: Opens a new word document to insert the desired files into.  After insertion,
        '   updates all fields, unlinks, and prompts the user to save in proper client name and
        '   file type format.

        Dim formType, formVersion, goFileLetter, goFileName As String
        Dim docName As String
        Dim insertFile1, insertFile2, insertFile3, insertFile4, insertFile5 As String
        Dim writer As String
        Dim dashPosition As Integer
        Dim clientName As String
        Dim goDoc, precedentDoc As Document
        Dim saveStatus As String
        Dim ppath, calcFile, ok, status As String
        goDoc = ActiveDocument
        On Error GoTo ErrorHandlerProtected

        ' setDefaultFormField update the form field value you are currently in
        status = updateAllBookmarks
        If (status = FUNCTION_ERROR) Then GoTo ErrorHandlerProtected

        ' Find information about the current go file.  Used to find the
        ' correct path to the calc file, and make decisions of how to differently
        ' update the file depending on type of law, ie F, RE, etc.  It is also
        ' used to suggest an appropriate file name when saving the merged file
        formType = ActiveDocument.FormFields("formType").TextInput.Default
        formVersion = ActiveDocument.FormFields("formVersion").TextInput.Default
        goFileLetter = Left(formType, 1)
        goFileName = goFileLetter & "Go" & formVersion
        docName = goDoc.name
        dashPosition = InStr(1, docName, "-")
        ' Client name here to include the folder type eg: Jones-W or Smith-F
        clientName = Left(docName, dashPosition + 1)

        ' Read the files that the user has inputted to be inserted/merged
        insertFile1 = trimWhiteSpace(goDoc.Bookmarks("insertFile1").Range)
        insertFile2 = trimWhiteSpace(goDoc.Bookmarks("insertFile2").Range)
        insertFile3 = trimWhiteSpace(goDoc.Bookmarks("insertFile3").Range)
        insertFile4 = trimWhiteSpace(goDoc.Bookmarks("insertFile4").Range)
        insertFile5 = trimWhiteSpace(goDoc.Bookmarks("insertFile5").Range)
        writer = trimWhiteSpace(goDoc.Bookmarks("writer").Range)
        unprotectDocument()
        On Error GoTo ErrorHandlerUnprotected
        ' find precedent path and calc file path

        ppath = getPrecedentPath()
        If (ppath = FUNCTION_ERROR) Then GoTo ErrorHandlerUnprotected
        ppath = ppath & PRECEDENT_FOLDER & "\"
        ppath = Replace(ppath, "\", "\\")
        calcFile = goFileLetter & "00-Calc.doc"

        ' Go to the third section of the document, the "document assembly section".
        ' In this section, add all of the includeText, Set ppath, etc values.
        ' First clear the whole section so don't have conflicting Sets, etc left
        ' over from previous builds.
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.EndKey(Unit:=wdStory, Extend:=wdExtend)
        Selection.Delete()

        ' Now that the section is cleared, go to the section & define ppath, pDesc, etc
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
            Text:="SET  PPath " & """" & ppath & """")

        ' insert/include the calc File or Function
        If (goFileLetter = "R") Then
            status = RECalcFunction()
            If (status = FUNCTION_ERROR) Then GoTo ErrorHandlerUnprotected
        Else
            ok = insertIncludeText(ppath, calcFile)
        End If

        ' insertFiles each file if does not contain default fill of xxx
        ok = insertIncludeText(ppath, insertFile1)
        ok = insertIncludeText(ppath, insertFile2)
        ok = insertIncludeText(ppath, insertFile3)
        ok = insertIncludeText(ppath, insertFile4)
        ok = insertIncludeText(ppath, insertFile5)

        If (goFileLetter = "W") Then
            insertHisHersWillInformation("Do not delete/modify the next 4 lines:  Info used to convert his to hers wills: ")
        End If

        ' Go to final section and construct the document there.  Update, and then
        ' copy and paste into a new file.  Ask the user for file saving Specifics
        ' insert the file stuff

        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.EndKey(Unit:=wdStory, Extend:=wdExtend)
        Options.DefaultHighlightColorIndex = wdNoHighlight
        Selection.Range.HighlightColorIndex = wdNoHighlight
        Selection.Copy()

        Documents.Add DocumentType:=wdNewBlankDocument
        precedentDoc = ActiveDocument
        Selection.Paste()
        Selection.WholeStory()
        Selection.fields.Unlink()
        showSpellErrors()

        Selection.HomeKey Unit:=wdStory

        Dim suggestName As String
        suggestName = suggestFileSaveName(insertFile1, insertFile2, insertFile3, insertFile4, insertFile5)
        saveStatus = saveDialog(clientName & "-" & suggestName, goDoc.path)
        If (saveStatus = FUNCTION_ERROR) Then GoTo ErrorHandlerUnprotected

        'Insert goHistory - only want to if the user saved the file
        If (saveStatus <> USER_EXIT) Then
            Dim files As String
            goDoc.Activate()
            Selection.HomeKey Unit:=wdStory
            Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=1, name:="")
            files = insertGoHistory(writer, insertFile1, insertFile2, insertFile3, insertFile4, insertFile5)
        End If

        goDoc.Activate()

        If (goFileLetter = "R") Then
            ActiveDocument.fields.ToggleShowCodes()
            ok = clearREFieldsAfterCalcFile()
            If (ok = FUNCTION_ERROR) Then GoTo ErrorHandlerUnprotected
            ActiveDocument.fields.ToggleShowCodes()
        End If

        ' clear the merged document in the 2nd section.
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.EndKey(Unit:=wdStory, Extend:=wdExtend)
        Selection.Cut()

        ClearClipboard()
        Selection.HomeKey Unit:=wdStory
        changeMode(CLIENT_MODE)
        protectDocument(wdAllowOnlyFormFields)
        precedentDoc.Activate()

        Exit Function

ErrorHandlerProtected:
        MsgBox("Error in updateAllGoForm protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)

        updateAllGoForm = FUNCTION_ERROR
        Exit Function

ErrorHandlerUnprotected:
        MsgBox("Error in updateAllGoForm unprotected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        goDoc.Activate()
        ' clear the 'assembly section' so looks nicer to user
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.EndKey(Unit:=wdStory, Extend:=wdExtend)
        Selection.Cut()
        goDoc.Activate()
        protectDocument(wdAllowOnlyFormFields)
        ClearClipboard()
        updateAllGoForm = FUNCTION_ERROR
    End Function

    Function getPrecedentPath() As String
        ' getPrecedentPath - Nov 2007 Andrea
        ' DESCRIPTION: This function returns the path where the precedents reside (thus precedentPath)
        '   differently depending on if the system is working on Vista or Windows XP.  The constants
        '   XP_PRECEDENT_LOCATION, etc are globals located in the module "DocumentConstruction"

        ' Set f based on vista or XP
        Dim f, ppath As String
        Dim errorString As String
        Dim folderExists As Boolean
        On Error GoTo ErrorHandler

        folderExists = checkDirectoryForFolder(XP_PRECEDENT_LOCATION, PRECEDENT_FOLDER)
        If (folderExists) Then
            f = XP_PRECEDENT_LOCATION
        Else
            folderExists = checkDirectoryForFolder(VISTA_PRECEDENT_LOCATION, PRECEDENT_FOLDER)
            If (folderExists) Then
                f = VISTA_PRECEDENT_LOCATION
            Else
                errorString = "CustomError: The folder " & PRECEDENT_FOLDER & " containing" & _
                " the precedents must reside in one of two locations.  " & _
                XP_PRECEDENT_LOCATION & " for Windows XP and " & VISTA_PRECEDENT_LOCATION & _
                " for Windows Vista"
                MsgBox(errorString)
                GoTo ErrorAfterMessage
            End If
        End If
        getPrecedentPath = f

        Exit Function
ErrorHandler:
        MsgBox("Error in getPrecedentPath: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
ErrorAfterMessage:
        getPrecedentPath = FUNCTION_ERROR
        Exit Function

    End Function

    Function insertIncludeText(ByVal ppath As String, ByVal insertFile As String) As String
        ' getInsertFiles - Nov 26_07 Andrea
        ' DESCRIPTION:  Inserts a reference to include the file located at input values
        '   insertFile1 and ppath where ever the cursor is currently located.  I.E. it is
        '   the responsibility of the calling function to determine where this is inserted.
        Dim temp As String
        Dim ansi As Integer
        On Error GoTo ErrorHandler

        ' Find ansi code of the first character in insertFile.  If it is 32, it is a
        ' white space and should not be included.
        ansi = Asc(insertFile)
        If (insertFile <> "xxx" And insertFile <> "" And ansi <> 32) Then
            temp = "Includetext " & """" & ppath & insertFile & """ " & """body"" "
            Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, Text:=temp)
        End If
        insertIncludeText = NO_ERROR

        Exit Function
ErrorHandler:
        MsgBox("Error in insertIncludeText: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        insertIncludeText = FUNCTION_ERROR
    End Function

    Function insertGoHistory(ByVal writer As String, ByVal insertFile1 As String, _
                        ByVal insertFile2 As String, ByVal insertFile3 As String, _
                        ByVal insertFile4 As String, ByVal insertFile5 As String) As String
        ' insertGoHistory - Nov 26_07 Andrea
        ' DESCRIPTION: Inserts the goFile insertion history. For example this function will
        '   insert into the goFile at the appropriate place (as determined by the calling
        '   function:  on December, 2007 the gofile was updated using the W20WillEpoaPerDirInv
        '   precedents by the author J/G-*
        Dim numLines As Long
        Dim line As String
        Dim compare As Integer
        Dim files, ok As String
        On Error GoTo ErrorHandler

        ' Go to the top of the document, and go to 2nd section
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToFirst, count:=2, name:="")
        ' Search for the "End" string
        Do While compare <> 0
            ' Search for the "End" string in the GoFile History Section
            Selection.Collapse Direction:=wdCollapseEnd
            numLines = Selection.MoveEnd(wdLine, 1)
            Selection.MoveLeft(Unit:=wdCharacter, count:=1, Extend:=wdExtend)
            line = Trim$(Selection.Text)
            Selection.MoveRight(Unit:=wdCharacter, count:=1, Extend:=wdExtend)
            compare = StrComp(line, "End")
        Loop

        ' Move to end of line, and to next:
        Selection.MoveDown(Unit:=wdLine, count:=1)

        ' Create String of information to be inserted in GoFile History, including
        ' writer, date, which files were updated
        files = ""
        files = concatInsertFiles(insertFile1, files)
        files = concatInsertFiles(insertFile2, files)
        files = concatInsertFiles(insertFile3, files)
        files = concatInsertFiles(insertFile4, files)
        files = concatInsertFiles(insertFile5, files)
        files = files & " -- Updated by " & writer & " On " & DateTime$
        Selection.TypeText Text:=files
        Selection.TypeParagraph()
        insertGoHistory = files

        Exit Function
ErrorHandler:
        MsgBox("Error in insertGoHistory: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        insertGoHistory = FUNCTION_ERROR
    End Function

    Function concatInsertFiles(ByVal fileAdd As String, ByVal allFiles As String) As String
        ' concatInsertFiles - Nov27_07 Andrea
        ' DESCRIPTION:  As used by insertGoHistory, if the inputString fileAdd is not the
        '   dummy value of "xxx", it returns a concatenated string of allFiles & fileAdd
        Dim ansi As Integer
        On Error GoTo ErrorHandler

        ' Find ansi code of the first character in file to add.  If it is 32, it is a
        ' white space and should not be included.
        ansi = Asc(fileAdd)
        If (fileAdd <> "xxx" And fileAdd <> "" And ansi <> 32) Then
            If (allFiles = "") Then
                concatInsertFiles = fileAdd
            Else
                concatInsertFiles = allFiles & " -- " & fileAdd
            End If
        Else
            concatInsertFiles = allFiles
        End If

        Exit Function
ErrorHandler:
        MsgBox("Error in concatInsertFiles: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        concatInsertFiles = FUNCTION_ERROR

    End Function


    Function suggestFileSaveName(ByVal insertFile1 As String, ByVal insertFile2 As String, _
                                 ByVal insertFile3 As String, ByVal insertFile4 As String, _
                                 ByVal insertFile5 As String) As String
        ' suggestFileSaveName - Nov27_07 Andrea
        ' DESCRIPTION:  As used By updateAllForm, combines all of the inserted file names together
        '   to create a suggestion for how the file should be saved

        Dim f As String
        On Error GoTo ErrorHandler

        f = removeInsertFileNumbersAndConcat(insertFile1, "")
        f = removeInsertFileNumbersAndConcat(insertFile2, f)
        f = removeInsertFileNumbersAndConcat(insertFile3, f)
        f = removeInsertFileNumbersAndConcat(insertFile4, f)
        f = removeInsertFileNumbersAndConcat(insertFile5, f)
        suggestFileSaveName = f

        Exit Function
ErrorHandler:
        MsgBox("Error in suggestFileSaveName: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        suggestFileSaveName = FUNCTION_ERROR

    End Function


    Function removeInsertFileNumbersAndConcat(ByVal fileAdd As String, _
                                              ByVal allFiles As String) As String
        ' removeInsertFileNumbersAndConcat - Nov27_07 Andrea
        ' DESCRIPTION: if the string is not "xxx", returns the fileAdd input without the leading
        '   numbers as found on filenames.  Ie takes "R45PrecedentAction" and retuns "PrecedentAction",
        '   or takes "xxx" and returns ""
        Dim ansi As Integer
        Dim cutOff As Integer

        On Error GoTo ErrorHandler

        ' Find ansi code of the first character in file to add.  If it is 32, it is a
        ' white space and should not be included.
        ansi = Asc(fileAdd)

        ' set variable cut off - for most of our precedents, cut off the last 3 letters (ie F45, etc).  If
        ' it is of type N (so probates, estates) the NC11, etc numbers mean something, so don't cut them off
        If (Left(fileAdd, 1) = "N") Then
            cutOff = 0
        Else
            cutOff = 3
        End If

        If (fileAdd <> "xxx" And fileAdd <> "" And ansi <> 32) Then
            If (allFiles = "") Then
                removeInsertFileNumbersAndConcat = Right(fileAdd, Len(fileAdd) - cutOff)
            Else
                removeInsertFileNumbersAndConcat = allFiles & "_" & Right(fileAdd, Len(fileAdd) - cutOff)
            End If
        Else
            removeInsertFileNumbersAndConcat = allFiles
        End If

        Exit Function
ErrorHandler:
        MsgBox("Error in removeInsertFileNumbersAndConcat: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        removeInsertFileNumbersAndConcat = FUNCTION_ERROR

    End Function


    Function checkDirectoryForFolder(directory As String, folder As String) As String
        ' DESCRIPTION:  The code from this function mainly follows the code found
        '   on the MSDN website under the "Visual Basic Reference, Dir Function".
        '   It is very similar to the last example given in the reference

        Dim readFolder
        On Error GoTo fcnErrorHandler

        'Initialize the return value as false so only if find set to true
        checkDirectoryForFolder = False
        readFolder = Dir(directory, vbDirectory)
        'Keep reading in the next folder, and check to see if folder matches
        Do While readFolder <> "" Or checkDirectoryForFolder = True
            If (readFolder = folder) Then
                checkDirectoryForFolder = True
            End If
            readFolder = Dir()              ' Get next entry.
        Loop
        Exit Function        ' If no error occurs, code will branch here
fcnErrorHandler:

        If (Err.Number = 52) Then
            'Here assume That the user will handle the error of the folder not existing
            Exit Function
        End If

    End Function


    Function trimWhiteSpace(ByVal inString As String)
        ' trimWhiteSpace - Jan14 2008 Andrea
        ' DESCRIPTION:  To replace the built in trim function - removes all trailing and leading
        '   white spaces/or 'other' formatting characters.

        Dim i As Integer
        Dim firstGood As Integer
        Dim lastGood As Integer
        Dim length As Integer
        Dim ansi As Integer
        Dim tempString As String
        Dim returnString As String
        On Error GoTo ErrorHandler

        'initialize.  If stay zero, know never found first/lastGood
        firstGood = 0
        lastGood = 0
        length = Len(inString)

        For i = 1 To length Step 1
            ' Find ansi code of the first character in file to add.  If it is 32, it is a
            ' white space and should not be included.
            tempString = Mid(inString, i, 1)
            ansi = Asc(tempString)
            If (ansi >= 33 And ansi <= 126) Then
                firstGood = i
                Exit For
            End If
        Next i

        For i = length To 1 Step -1
            ' Find ansi code of the first character in file to add.  If it is 32, it is a
            ' white space and should not be included.
            tempString = Mid(inString, i, 1)
            ansi = Asc(tempString)
            If (ansi >= 33 And ansi <= 126) Then
                lastGood = i
                Exit For
            End If
        Next i

        If (firstGood > lastGood) Then
            GoTo ErrorHandler
        End If
        If (firstGood = 0 And lastGood = 0) Then
            returnString = "     "
        Else
            'if found usual case with acceptable within
            returnString = Mid(inString, firstGood, lastGood - firstGood + 1)
        End If

        trimWhiteSpace = returnString
        Exit Function

ErrorHandler:
        MsgBox("Error in trimWhiteSpace: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        trimWhiteSpace = FUNCTION_ERROR

    End Function

    Function ClearClipboard()

        ' clearClipboard - Andrea 2008Jan
        ' DESCRIPTION:  Clears the clipboard to avoid the annoying message when you leave word of:
        '   "You have a large amount on the clipboard... would you like to save it for future use?"
        '   Following lines writes an empty object, and writes the emptiness to the clipboard to clear

        'Dim oData   As New DataObject
        'oData.SetText Text:=Empty
        'oData.PutInClipboard

        'I found this on the net on how to clear clipboard... the other method was causing errors
        'CommandBars.FindControl(ID:=3634).Execute
        'CommandBars("Clipboard").Controls("Clear Clipboard").Execute

        Dim tempData As DataObject
        Dim clipboardString As String
        tempData = New DataObject
        ' Clears the clipboard
        tempData.SetText ""
        tempData.PutInClipboard()

    End Function



    Function getProtection() As String
        ' getProtection - Andrea April 2008
        ' DESCRIPTION: Returns the protection of the active Document
        getProtection = ActiveDocument.protectionType

    End Function

    Function getMode() As String
        ' getMode - Andrea 2008Jan
        ' DESCRIPTION: This function returns the current mode of the system.  It
        '   needs to be a function because we also need to support previous versions of
        '   the system.  Eventually, will be able to drop this function & just get mode
        '   from formfield, but until then...

        Dim formVersion As String
        Dim currentMode As String
        Dim protection As String
        Dim msg As String

        On Error GoTo ErrorHandler
        formVersion = ActiveDocument.FormFields("formVersion").TextInput.Default
        formVersion = Left(formVersion, 1)
        protection = ActiveDocument.protectionType

        If (formVersion = "1") Then
            ' then want to set MODE manually
            If (protection = wdAllowOnlyReading) Then
                currentMode = START_MODE
            ElseIf (protection = wdAllowOnlyFormFields) Then
                currentMode = CLIENT_MODE
            End If
            'note: know could never get into HELP or DEFAULT mode as
            'the user will no longer have access to the start up blank
            'r00-goForm1.0 anymore.
        Else
            ' read in mode from formField
            currentMode = ActiveDocument.FormFields("goMode").result
            If (currentMode = CLIENT_MODE And protection <> wdAllowOnlyFormFields) Then
                GoTo ErrorWrongModeProtection
            ElseIf (currentMode = START_MODE And protection <> wdAllowOnlyReading) Then
                GoTo ErrorWrongModeProtection
            ElseIf (currentMode = HELP_MODE And protection <> wdAllowOnlyFormFields) Then
                GoTo ErrorWrongModeProtection
            ElseIf (currentMode = DEFAULT_MODE And protection <> wdAllowOnlyFormFields) Then
                GoTo ErrorWrongModeProtection
            End If
        End If
        getMode = currentMode

        Exit Function
ErrorHandler:
        MsgBox("Error in getMode: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        getMode = FUNCTION_ERROR

        Exit Function
ErrorWrongModeProtection:
        MsgBox("Error in getMode: " & "the protection type does not match the mode type")
        getMode = FUNCTION_ERROR
    End Function

    Function changeMode(ByVal desiredMode As String) As String
        ' getMode - Andrea 2008Jan
        ' DESCRIPTION: This function Changes the current mode of the system.  It
        '   needs to be a function because we also need to support previous versions of
        '   the system.  Eventually, will be able to drop this function & just get mode
        '   from formfield, but until then...
        ' NOTE: Can't change mode when document is locked for editing - ie readonly protection

        Dim formVersion As String
        Dim msg As String
        Dim protection As String
        On Error GoTo ErrorHandler

        ' ensure protection is none, as can't change mode otherwise
        protection = ActiveDocument.protectionType
        If (protection = wdAllowOnlyReading Or protection = wdAllowOnlyFormFields) Then
            MsgBox("Error:  Function changeMode can only be called on an unprotected file")
            GoTo ErrorHandlerAfterMessage
        End If

        ' ensure mode passed is one of the options:
        Select Case desiredMode
            Case START_MODE, CLIENT_MODE, HELP_MODE, DEFAULT_MODE, ERROR_MODE
            Case Else
                msg = "The function changeMode must receive a valid input string" & _
                      ".  The mode: " & desiredMode & " is not a valid choice. " & _
                      " Exiting...."
                GoTo ErrorHandlerAfterMessage
        End Select

        formVersion = ActiveDocument.FormFields("formVersion").TextInput.Default
        formVersion = Left(formVersion, 1)
        If (formVersion = "1") Then
            ' then want to do nothing
        Else
            ' change mode variable
            ActiveDocument.FormFields("goMode").TextInput.Default = desiredMode
            ActiveDocument.FormFields("goMode").result = desiredMode
            ActiveDocument.Bookmarks("goMode").Range.fields(1).Update()
        End If
        changeMode = NO_ERROR

        Exit Function
ErrorHandler:
        MsgBox("Error in changeMode: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
ErrorHandlerAfterMessage:
        changeMode = FUNCTION_ERROR
    End Function

    Sub updateForExcelGoImport()
        ' updateaAllGoForm - Nov 20_07 Andrea
        ' DESCRIPTION: Opens a new word document to insert the desired files into.  After insertion,
        '   updates all fields, unlinks, and prompts the user to save in proper client name and
        '   file type format.

        Dim formType, formVersion, goFileLetter, goFileName As String
        Dim docName As String
        Dim goDoc As Document
        Dim ppath, calcFile, ok, status As String
        goDoc = ActiveDocument
        Dim protection As String
        Dim currentMode As String
        On Error GoTo ErrorHandlerProtected

        protection = ActiveDocument.protectionType

        If (protection <> wdAllowOnlyReading And protection <> wdAllowOnlyFormFields) Then
            GoTo PressInError
        End If

        currentMode = getMode()

        If (currentMode <> CLIENT_MODE) Then
            GoTo PressInError
        End If

        status = updateAllBookmarks
        If (status = FUNCTION_ERROR) Then GoTo ErrorHandlerProtected

        ' Find information about the current go file.  Used to find the
        ' correct path to the calc file, and make decisions of how to differently
        ' update the file depending on type of law, ie F, RE, etc.  It is also
        ' used to suggest an appropriate file name when saving the merged file
        unprotectDocument()
        On Error GoTo ErrorHandlerUnprotected
        ' find precedent path and calc file path
        formType = ActiveDocument.FormFields("formType").TextInput.Default
        formVersion = ActiveDocument.FormFields("formVersion").TextInput.Default
        goFileLetter = Left(formType, 1)
        goFileName = goFileLetter & "Go" & formVersion

        ppath = getPrecedentPath()
        If (ppath = FUNCTION_ERROR) Then GoTo ErrorHandlerUnprotected
        ppath = ppath & PRECEDENT_FOLDER & "\"
        ppath = Replace(ppath, "\", "\\")
        calcFile = goFileLetter & "00-Calc.doc"

        ' Go to the third section of the document, the "document assembly section".
        ' In this section, add all of the includeText, Set ppath, etc values.
        ' First clear the whole section so don't have conflicting Sets, etc left
        ' over from previous builds.
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.EndKey(Unit:=wdStory, Extend:=wdExtend)
        Selection.Delete()

        'Now go to the now cleared section & add the definition to ppath
        Selection.HomeKey Unit:=wdStory
        Selection.GoTo(What:=wdGoToSection, Which:=wdGoToNext, count:=2, name:="")
        Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
            Text:="SET  PPath " & """" & ppath & """")

        ' insert/include the calc File or Function - include definitions for pDesc, etc
        If (goFileLetter = "R") Then
            status = RECalcFunction()
            If (status = FUNCTION_ERROR) Then GoTo ErrorHandlerUnprotected
        Else
            ok = insertIncludeText(ppath, calcFile)
        End If

        updateAllBookmarks()

        Selection.HomeKey Unit:=wdStory
        changeMode(CLIENT_MODE)
        protectDocument(wdAllowOnlyFormFields)

        Exit Sub

ErrorHandlerProtected:
        MsgBox("Error in updateForExcelGoImport protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        updateAllGoForm = FUNCTION_ERROR
        Exit Sub

ErrorHandlerUnprotected:
        MsgBox("Error in updateForExcelGoImport unprotected: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        protectDocument(wdAllowOnlyFormFields)
        updateAllGoForm = FUNCTION_ERROR
        Exit Sub

PressInError:
        MsgBox("The goFile you have selected is NOT of the new version type. " & _
               "Exiting.  No action taken, no GoFile information imported")
        Exit Sub
    End Sub


End Module