Module DocumentConstruction

Option Explicit On

    '***********************************************************
    ' MODULE DESCRIPTION:  DocumentConstruction.  This module contains all macros and functions that are
    '   necessary for the OLD non form document construction package such as go files and precedents.
    '   This package was overhauled when the new PC's with vista were purchased for the office.  The
    '   new construction package using form fields does not call any functions from this module.  It
    '   only uses the global constants that are declared at the top of the module.  These functions
    '   are slightly messy, as it was a difficult job to fix such a non uniform system (ie to
    '   update a family go file is different than to update a real estate as includes, etc are
    '   located in different places in the goFile itself.
    '***********************************************************

    Function documentConstructionAction()
        ' NOT IN USE.  CAN NOT REALIABLY ENOUGH DISCERN BETWEEN EMPTY & CLIENT GO FILES
        '  documentConstructionAction - Andrea oct 2007
        ' DESCRIPTION:  This function combines all actions that a user of the document construction
        '   package would use into one function.  These functions include UpdateAll, Save as Go,
        '   unlinkAllSaveAs, and no action if pressed in error.  This function decides what the
        '   user would like to do based on document properties - what is in the active Docuemnt
        '   name, what kind of protection, etc.
        ' SHORTCUT: Not yet assigned on the idea that going on the file name is not enough to
        '   distinguish what action the user would like to take.

        Dim fileName As String
        Dim findPositionZero, findPositionDash, findPositionGo As Integer
        Dim protectionType As String
        Dim errorMsg As String

        protectionType = ActiveDocument.protectionType

        If (protectionType = wdAllowOnlyReading Or protectionType = wdAllowOnlyFormFields) Then
            errorMsg = "This file is of wrong Protection type to use the old GoFile " & _
            " macros.  If you are in new GoForms consider using F3 instead.  Exiting, no Action performed."
            Exit Function
        End If

        fileName = ActiveDocument.name
        findPositionZero = InStr(1, fileName, "00-")
        findPositionDash = InStr(1, fileName, "-")
        findPositionGo = InStr(1, fileName, "Go")

        If findPositionZero <> 0 And findPositionDash <> 0 And findPositionGo <> 0 Then
            ' We know we are in a blank go file as have a 00-, ie W00-Go
            Application.Run MacroName:="SaveAsGo"
        ElseIf findPositionZero = 0 And findPositionDash <> 0 And findPositionGo <> 0 Then
            ' We know there is no -00 in the file name, and has a go.  We are in a client go file.
            Application.Run MacroName:="UpdateAll"
            Application.Run MacroName:="UnlinkAllSaveAs"
        Else
            msgBox("Custom Error:  This file is NOT a goFile.  SaveAsGoFile and updateGoFile " & _
            "only work on a file that is of type goFile, and where the document is named as such.")
        End If

    End Function

    Sub SaveAsGo()
        ' SaveAsGo - 05June15 - Grant and James Dukeshire, Minor Mod format only Andrea, Nov 2007
        ' DESCRIPTION:  This function is called by the user everytime a go file is
        '   opened.  As the go file is locked for editing, the user must 'save as'
        '   a copy of the file, to the current path, and use an inputted clientName to help
        '   streamline the saving and file naming process
        ' KEY SHORTCUT: Ctrl-Shift-F12 .Reason piggyback UnlinkAllSave

        Dim dashPosition As Integer '  from W00-WGo1 (firstDash)
        Dim clientName As String       '  eg:  Smith- to replace W00-
        Dim goFileName As String       '  "CGo1", "DGo1", ....   "WGo1", etc, SO TYPE, Wills, Sales, etc
        Dim docName As String       '  path-workFolder-clientName-docName.doc
        Dim StrClean1 As Object
        Dim StrClean2 As Object
        Dim StrClean3 As Object
        Dim protectionType As String
        Dim errorMsg As String

        protectionType = ActiveDocument.protectionType
        If (protectionType = wdAllowOnlyReading Or protectionType = wdAllowOnlyFormFields) Then
            errorMsg = "This file is of wrong Protection type to use the old GoFile " & _
            " macros.  If you are in new GoForms consider using F3 instead.  Exiting, no Action performed."
            Exit Sub
        End If

        On Error GoTo ErrorHandlerSaveAsGo
        ' Assume ActiveDoc is of form W00-WGo1 & needs to change to Smith-WGo1
        dashPosition = InStr(1, ActiveDocument.name, "-")            'first "-" in doc
        goFileName = Mid(ActiveDocument.name, dashPosition + 1, 4)   'WGo1

        ' Check if FileName looks like Smith-WGo1 & warn that "SmithGo" already made
        If Mid(ActiveDocument.name, dashPosition - 2, 3) <> "00-" Then
            msgBox("Warning = You are making a duplicate of a GoFile ALREADY named")
        End If

        ' Ask user for clientName, trim blanks, check for blank & null conditions & "non-doc"
        clientName = InputBox("Enter Client LastNameID to save", , "LastNameID") 'Smith
        clientName = Trim(clientName)
        StrClean1 = InStr(1, clientName, "/")
        StrClean2 = InStr(1, clientName, "\")
        StrClean3 = InStr(1, clientName, "-")
        If StrClean1 <> 0 Or StrClean2 <> 0 Or StrClean3 <> 0 Then
            msgBox("Error = LastNameID cannot have a '/' or '\'or '-'.  File not saved")
            Exit Sub
        End If
        If clientName = "" Or IsNull(clientName) Then Exit Sub

        ' Following Opens the normal SaveAs window and suggests a default input name
        ' being the default path where the InitialFileName GoFile resides that was opened
        ' If a BOGUS path is given it will go back to the last path that (default path)
        ' the text after the last "\" in the path is what is used as the default name
        Dim dlgSaveAs As FileDialog
        dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
        dlgSaveAs.InitialFileName = ActiveDocument.path & "\" & clientName & "-" & goFileName
        dlgSaveAs.InitialView = msoFileDialogViewDetails    ' shows details
        dlgSaveAs.title = "SaveAsGo Macro" & " - SaveAsGo"
        dlgSaveAs.ButtonName = "Save GoFile"
        If dlgSaveAs.Show = -1 Then dlgSaveAs.Execute()
        Exit Sub

ErrorHandlerSaveAsGo:
        Dim msg As String
        msg = "An error occured! " & "Error number: " & Err.Number & "Error message: " & Err.Description
        msgBox(msg, vbOKOnly + vbCritical)
    End Sub

    Sub UpdateAll()
        ' UpdateAll - 05June30 - James Dukeshire, Major Modification Nov 2007 by Andrea
        ' DESCRIPTION:  This function is called by the user after all data has been inputted to the go
        '   file, ie the fields have their appropriate values for the client.  The fields are updated,
        '   and the path where the precedents reside (PPath) is determined.  PPath is then used within
        '   the gofile to insert the appropriate precedents as desired.
        ' KEY SHORTCUT: F9.  Reason for ALL fields as is the key MS used for a Single field

        ' Error Checking:  Ensure all bookmarks exist that the macro needs to work with
        Dim CalcTypeExist, FullCalcPathExist, PPathExist As Boolean
        On Error GoTo UpdateAllErrorHandler
        CalcTypeExist = ActiveDocument.Bookmarks.Exists("CalcType")
        FullCalcPathExist = ActiveDocument.Bookmarks.Exists("FullCalcPath")
        PPathExist = ActiveDocument.Bookmarks.Exists("PPath")

        ' Checking the file type, as determined by the doc's filename.  Ie W00-Go is a Will file
        Dim firstLetter, goFileName, calcString As String
        Dim isC, isD, isE, isF, isR, isW As Boolean
        Dim dashPosition As Integer
        isC = False
        isD = False
        isE = False
        isF = False
        isR = False
        isW = False
        calcString = "xxxx"

        dashPosition = InStr(1, ActiveDocument.name, "-")            'first "-" in doc
        goFileName = Mid(ActiveDocument.name, dashPosition + 1, 4)   'WGo1 or FGo1
        firstLetter = Left$(goFileName, 1)

        Select Case firstLetter
            Case "C"
                isC = True
            Case "D"
                isD = True
                calcString = "D00-Calc"
            Case "E"
                isE = True
                calcString = "E00-Calc"
            Case "F"
                isF = True
                calcString = "F00-Calc"
            Case "R"
                isR = True
            Case "W"
                isW = True
                calcString = "W00-Calc"
            Case Else
                msgBox("Custom Error: The file extension must be in Go file format ie Smith-WGo1.  The document must be saved in this format to use this macro")
                Exit Sub
        End Select

        ' Set f based on vista or XP
        Dim f, errorString As String
        Dim folderExists As Boolean

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
                msgBox(errorString)
                f = " "
                Exit Sub
            End If
        End If

        'WINDOWS_7_PRECEDENT_LOCATION
        folderExists = checkDirectoryForFolder(WINDOWS_7_PRECEDENT_LOCATION, PRECEDENT_FOLDER)
        If (folderExists) Then
            f = WINDOWS_7_PRECEDENT_LOCATION
        Else
            folderExists = checkDirectoryForFolder(VISTA_PRECEDENT_LOCATION, PRECEDENT_FOLDER)
            If (folderExists) Then
                f = VISTA_PRECEDENT_LOCATION
            Else
                folderExists = checkDirectoryForFolder(XP_PRECEDENT_LOCATION, PRECEDENT_FOLDER)
                If (folderExists) Then
                    f = XP_PRECEDENT_LOCATION

                Else

                    errorString = "CustomError: The folder " & PRECEDENT_FOLDER & " containing" & _
                    " the precedents must reside in one of 3 locations.  " & _
                    WINDOWS_7_PRECEDENT_LOCATION & " for windows 7 or " & VISTA_PRECEDENT_LOCATION & _
                    " for Windows Vista or " & XP_PRECEDENT_LOCATION & " for Windows XP"
                    msgBox(errorString)
                    f = " "
                    Exit Sub
                End If
            End If
        End If

        ' Using the path of zPrecedents we just found in f, prepare the values of
        ' stringPPath and stringFullCalcPath(NOT bookmark PPath & bookmark FullCalcPath).
        ' The Values of these strings will be written into the appropriate bookmarks
        ' in different ways depending on the type of Go file we are
        ' using such as Family (isF), Wills(isW), etc.
        Dim string1, string2, string3, stringPPath, stringFullCalcPath As String

        ' Create and Set the value of the PPath Bookmark
        string1 = f & PRECEDENT_FOLDER & "\"
        string2 = "SET  PPath " & """" & string1 & """"
        stringPPath = Replace(string2, "\", "\\")

        If isC Or isR Or isW Then
            If CalcTypeExist = False And FullCalcPathExist = False And isW = False Then
                msgBox("Custom Error: The C, R, or W Go file is not in the correct format to be updated.  See Macro for more details.  Exit.  File not updated")
                Exit Sub
            ElseIf CalcTypeExist = False And FullCalcPathExist = False And isW = True Then
                ' This line came from a specialized will go file that is in existance.
                ' This has a hardCode WRONG include statement, and no definition of calcType, so
                ' we treat it the same as a D, E, or F file, thus go to the tag insertIncludes
                If FullCalcPathExist = False Then
                    GoTo insertIncludes
                End If
            End If

            ' Set PPath and FullCalcPath
            InsertPPathSetField(stringPPath)
            string3 = string1 & ActiveDocument.FormFields("CalcType").TextInput.Default & ".doc"
            stringFullCalcPath = Replace(string3, "\", "\\")
            ActiveDocument.FormFields("FullCalcPath").TextInput.Default = stringFullCalcPath

            UpdateFields()


        ElseIf isD Or isE Or isF Then
insertIncludes:
            If CalcTypeExist = True Or FullCalcPathExist = True Then
                msgBox("Custom Error1: The D, E, or F Go file is not in the correct format to be updated.  See Macro for more details.  Exit.  File not updated")
                Exit Sub
            End If

            InsertPPathSetField(stringPPath)
            CutFirstIncludeText()

            ' Create & insert the String/field to correctly insert the calc file
            Dim stringInsertLinkToCalcFile As String
            string3 = string1 & calcString & ".doc"
            stringFullCalcPath = Replace(string3, "\", "\\")
            stringInsertLinkToCalcFile = "IncludeText " & """" & stringFullCalcPath & """" & """" & "body" & """"
            With ActiveDocument.Bookmarks
                .Add(Range:=Selection.Range, name:="InsertLinkToCalcFile")
            End With
            Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
                Text:=stringInsertLinkToCalcFile)
            UpdateFields()

        Else
            msgBox("Custom Error2: Logic Error.  This Loop should Never Get here.")
            Exit Sub
        End If
        ' Delete the instance of PPath bookmark that was inserted at the top of
        ' the document as a work around.  Must toggle codes back to the goFile view in
        ' order to delete the field instance. (Toggling happened in previous two if loops)
        ActiveDocument.fields.ToggleShowCodes()
        Selection.HomeKey Unit:=wdStory
        Selection.Delete(Unit:=wdCharacter, count:=1)
        ActiveDocument.fields.ToggleShowCodes()
        Exit Sub

UpdateAllErrorHandler:
        Dim msg As String
        msg = "An error occured! " & "Error number: " & Err.Number & "Error message: " & Err.Description
        msgBox(msg, vbOKOnly + vbCritical)
    End Sub

    Function CutFirstIncludeText()
        ' CutFirstIncludeText - Andrea oct 2007
        ' DESCRIPTION: used within the function update all to find the first 'IncludeText' reference, and delete it
        '   out.  This was neccesary to clean up old goFiles that were hardwired to a specific location in the
        '   file, and now is completely wrong.

        Selection.HomeKey Unit:=wdStory
        Selection.Find.ClearFormatting()
        With Selection.Find
            .Text = "IncludeText"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

        ' Have looked for it.  If there, delete it, else just return out
        If Selection.Find.Execute = True Then
            ' Cut it out
            Selection.Find.Execute()
            Selection.MoveLeft(Unit:=wdCharacter, count:=3)
            Selection.NextField.Select()
            Selection.Cut()
        End If

    End Function

    Function InsertPPathSetField(istringPPath)
        ' InsertPPathSetField - Andrea Oct 2007
        ' DESCRIPTION: Inserts a correct hardwired command to define PPath at the top of the goFile
        Selection.HomeKey Unit:=wdStory
        With ActiveDocument.Bookmarks
            .Add(Range:=Selection.Range, name:="PPath")
        End With
        Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, _
            Text:=istringPPath)
    End Function

    Function UpdateFields()
        ' UpdateFields - Andrea Oct 2007
        ' DESCRIPTION:  Once the fields are updated, we want to toggle the view of the field codes
        '   twice to 'kick' the system when we update the document.  'Kicking' is remanant of OLD
        '   programming of James'
        ActiveDocument.fields.Update()
        ActiveDocument.fields.ToggleShowCodes()
        ActiveDocument.fields.ToggleShowCodes()
        ActiveWindow.View.ShowFieldCodes = False
    End Function
    Sub UnlinkAllSaveAs()
        ' UnlinkAllSave - 05June30 - Grant & James Dukeshire, Mod minor formatting only Nov 07 Andrea
        ' Mod 04May30 for Saving per "-" from GoFiles - GD
        ' DESCRIPTION: Unlinks the field codes within a document to be formatted text
        '   and saves the file to a new name.  The user is prompted with
        '   a save path to Family-RealEstate-Wills folders based on where the GoFile is
        ' KEY SHORTCUT: Ctrl+F12. Reason as not to make TOO easy

        ' Will not work if GoFile locked - so ONLY for SmithGo's, not W00-Go1
        ' Checks if already protected & if so unprotects so macro can proceed
        ' Carefull with Protected Files!  Don't write over the go file while testing
        On Error GoTo ErrorHandlerUnlinkAllSaveAs

        Dim protectionType As String
        Dim errorMsg As String

        protectionType = ActiveDocument.protectionType

        If (protectionType = wdAllowOnlyReading Or protectionType = wdAllowOnlyFormFields) Then
            errorMsg = "This file is of wrong Protection type to use the old GoFile " & _
            " macros.  If you are in new GoForms consider using F3 instead.  Exiting, no Action performed."
            Exit Sub
        End If


        If ActiveDocument.protectionType <> wdNoProtection Then ActiveDocument.unprotect()

        Dim dashPos As Integer          ' from Smith-WGo1
        Dim clientName As String        ' eg:  Smith-W
        Dim docName As String           ' path-workFolder-clientName-docName.doc
        Dim docSavePath As String       ' Path to which Document is to be saved
        Dim StrClean1 As Object
        Dim StrClean2 As Object

        ' Assume the ActiveDoc is of form Smith-WGo1 (as know have run update all macro)
        ' Use Smith-W as prefix for new fileName (just add "-Will-his" to it)
        ' Then combine gstrHomeFolder & workFolder & prefix & fileName

        ' Obtains client name from the Gofile name to be used for variable docSavePath
        ' Client name here to include the folder type eg: Jones-W or Smith-F
        dashPos = InStr(1, ActiveDocument.name, "-")          ' Finds position of 1st "-"
        clientName = Left(ActiveDocument.name, dashPos + 1)   ' ie:  Jones-W

        ' Would like to call SaveAsGo() if FileName looks like W00-WGo1
        If Mid(ActiveDocument.name, dashPos - 2, 3) = "00-" Then
            msgBox("Error = GoFile NOT named - should have used Ctrl-Shift-F12. File not saved")
            Exit Sub
        End If

        ' Ask user for docName, trim blanks, check for blank & null conditions & "non-doc"
        docName = InputBox("Enter Document Name to save", , "Document Name")
        docName = Trim(docName)

        ' Following to protect document from error that occurs when "/" or "\" is put in
        StrClean1 = InStr(1, docName, "/")
        StrClean2 = InStr(1, docName, "\")
        If StrClean1 <> 0 Or StrClean2 <> 0 Then
            msgBox("Error = DocName cannot have a '/' or '\'.  File not saved")
            Exit Sub
        End If

        ' Checks for blank or null entries for document name
        If docName = "" Or IsNull(docName) Then Exit Sub

        ' Add ".doc" to end if not already there so it will open easily with Word
        If Right(docName, 4) <> ".doc" Then docName = docName & ".doc"

        ' docSavePath is path by which document is to be saved
        docSavePath = ActiveDocument.path & "\" & clientName & "-" & docName

        Dim dlgSaveAs As FileDialog
        ' Makes fileDialog SaveAs box that unlinks & does other neccesary functions
        dlgSaveAs = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
        dlgSaveAs.InitialFileName = docSavePath
        dlgSaveAs.title = "UnlinkAllSave"
        dlgSaveAs.InitialView = msoFileDialogViewDetails        ' shows details
        dlgSaveAs.ButtonName = "Create Doc"
        ' dlgSaveAs.ButtonName = "UnLink & Save Doc"

        ' Ensuring no saving over file of same name
        If dlgSaveAs.Show = -1 Then
            ' Updates before unlinking, in case user forgets (usually takes another 3-5 seconds)
            ' DO NOT update AFTER removing highlights -update will return some of the highlights
            ' ActiveDocument.Fields.Update
            ActiveDocument.Save()
            ActiveDocument.fields.Unlink()
            ' Ensures that ALL spelling mistakes detected in document
            ActiveDocument.Range.NoProofing = False
            ' Removes highlights for entire document
            ActiveDocument.Range.HighlightColorIndex = wdNoHighlight
            dlgSaveAs.Execute()
        End If
        Exit Sub

ErrorHandlerUnlinkAllSaveAs:
        Dim msg As String
        msg = "An error occured! " & "Error number: " & Err.Number & "Error message: " & Err.Description
        msgBox(msg, vbOKOnly + vbCritical)
    End Sub
End Module