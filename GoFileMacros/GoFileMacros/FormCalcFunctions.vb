Module FormCalcFunctions


Option Explicit
    ' Module: FormDocumentConstruction.  Contains all macros related
    ' only to the Real Estate section of law (RE)

    Function RECalcFunction() As String
        'This is a function to be called instead of including the RE calc file.
        ' If all the functionality is in here, is easier to debug, can stop &
        ' see what values things are getting, etc.

        ' Need to ensure that calling function is in the 'merging' section where
        ' we will insert the proper set calls.
        ' Instead of including a if, else setting calc file, will only write to the
        ' section what to set based on the if elses here.

        ' also need to ensure that are not using fields ie (SET var1 "LKJ") if they
        ' are already used in a form field -- The SW can only have it defined in one
        ' place & it corrupts the bookmark if done the other way.

        Dim re2 As String
        Dim legal1 As String
        Dim uAd1, uAd2, xAd1, xAd2 As String
        Dim uDesc, xDesc As String
        Dim ad1, ad2, ad3 As String
        Dim u1, u2, x1, x2 As String
        Dim v1, v2, p1, p2 As String
        Dim uv As String
        Dim v, p As String
        Dim vAd1, vAd2, pAd1, pAd2 As String
        Dim vDesc, pDesc As String
        Dim u1Spouse As String
        Dim aMCo, nMCo As String

        Dim u, x As String
        Dim setString1, setString2, setString3, setString4 As String
        Dim setString5, setString6, setString7, setString8 As String
        Dim setString9, setString10, setString11, setString12 As String
        Dim msg As String
        Dim ok As String
        Dim temp As String
        On Error GoTo ErrorHandler

        ' Read everything from file need to make decisions
        ' Want to make sure all of data/bkmarks are in the format that have whitespace
        ' trimmed off - that the result is what has been saved in the default

        re2 = trimWhiteSpace(ActiveDocument.FormFields("re2").TextInput.Default)
        legal1 = trimWhiteSpace(ActiveDocument.FormFields("legal1").TextInput.Default)
        uAd1 = trimWhiteSpace(ActiveDocument.FormFields("uAd1").TextInput.Default)
        uAd2 = trimWhiteSpace(ActiveDocument.FormFields("uAd2").TextInput.Default)
        xAd1 = trimWhiteSpace(ActiveDocument.FormFields("xAd1").TextInput.Default)
        xAd2 = trimWhiteSpace(ActiveDocument.FormFields("xAd2").TextInput.Default)
        uDesc = trimWhiteSpace(ActiveDocument.FormFields("uDesc").TextInput.Default)
        xDesc = trimWhiteSpace(ActiveDocument.FormFields("xDesc").TextInput.Default)

        ad1 = trimWhiteSpace(ActiveDocument.FormFields("ad1").TextInput.Default)
        ad2 = trimWhiteSpace(ActiveDocument.FormFields("ad2").TextInput.Default)
        ad3 = trimWhiteSpace(ActiveDocument.FormFields("ad3").TextInput.Default)
        u1 = trimWhiteSpace(ActiveDocument.FormFields("u1").TextInput.Default)
        u2 = trimWhiteSpace(ActiveDocument.FormFields("u2").TextInput.Default)
        x1 = trimWhiteSpace(ActiveDocument.FormFields("x1").TextInput.Default)
        x2 = trimWhiteSpace(ActiveDocument.FormFields("x2").TextInput.Default)
        uv = trimWhiteSpace(ActiveDocument.FormFields("uv").TextInput.Default)
        u1Spouse = trimWhiteSpace(ActiveDocument.FormFields("u1Spouse").TextInput.Default)
        aMCo = trimWhiteSpace(ActiveDocument.FormFields("aMCo").TextInput.Default)
        nMCo = trimWhiteSpace(ActiveDocument.FormFields("nMCo").TextInput.Default)

        'have legacy if re2, uad1, xad1, udesc, xdesc are empty.  Note each places the
        ' calculated values of re2, etc in the formfield, and then writes to a bookmark
        ' so that in later steps we can determine if the user entered this, or the program
        ' calculated the contents of re2, etc.
        If (re2 = "     " Or re2 = "") Then
            setString1 = ad1 & ", " & ad2 & " - " & legal1
            ActiveDocument.FormFields("re2").result = setString1
            re2 = setString1
            setString1 = "SET re2Delete " & """" & "True" & """"
            addSetField(setString1)
        End If

        If (uAd1 = "     " Or uAd1 = "") Then
            setString2 = ad1
            uAd1 = setString2
            ActiveDocument.FormFields("uAd1").result = setString2
            setString3 = ad2 & " " & ad3
            ActiveDocument.FormFields("uAd2").result = setString3
            uAd2 = setString3
            setString2 = "SET uAd1Delete " & """" & "True" & """"
            setString3 = "SET uAd2Delete " & """" & "True" & """"
            addSetField(setString2)
            addSetField(setString3)
        End If

        If (xAd1 = "     " Or xAd1 = "") Then
            setString4 = ad1
            ActiveDocument.FormFields("xAd1").result = setString4
            xAd1 = setString4
            setString5 = ad2 & "  " & ad3
            ActiveDocument.FormFields("xAd2").result = setString5
            xAd2 = setString5
            setString4 = "SET xAd1Delete " & """" & "True" & """"
            setString5 = "SET xAd2Delete " & """" & "True" & """"
            addSetField(setString4)
            addSetField(setString5)
        End If
        If (uDesc = "     " Or uDesc = "") Then
            temp = getDescriptionFromBookmarks("u1", "u2", "uAd1", "uAd2")
            ActiveDocument.FormFields("uDesc").result = temp
            uDesc = temp
            temp = "SET uDescDelete " & """" & "True" & """"
            addSetField(temp)
        End If

        If (xDesc = "     " Or xDesc = "") Then
            temp = getDescriptionFromBookmarks("x1", "x2", "xAd1", "xAd2")
            ActiveDocument.FormFields("xDesc").result = temp
            xDesc = temp
            temp = "SET xDescDelete " & """" & "True" & """"
            addSetField(temp)
        End If

        ' do combining sets - for u, x - here must use set as no bkm exist here
        setString1 = "SET u " & """" & u1
        If (u2 = "     " Or u2 = "") Then
            setString1 = setString1 & """"
            u = u1
        Else
            setString1 = setString1 & " and " & u2 & """"
            u = u1 & " and " & u2
        End If
        addSetField(setString1)

        setString2 = "SET x " & """" & x1
        If (x2 = "     " Or x2 = "") Then
            setString2 = setString2 & """"
            x = x1
        Else
            setString2 = setString2 & " and " & x2 & """"
            x = x1 & " and " & x2
        End If
        addSetField(setString2)

        ' do setting of v, p, etc based on the value of uv - perhaps check value
        ' of uv is an acceptable value - should this be a dropdown?
        If (uv <> "y" And uv <> "n") Then
            msg = "Warning:  The input to uv must be either y or n.  "
            msgBox msg
        End If

        If (uv = "y") Then
            setString1 = "SET v1 " & """" & u1 & """"
            setString2 = "SET v2 " & """" & u2 & """"
            setString3 = "SET p1 " & """" & x1 & """"
            setString4 = "SET p2 " & """" & x2 & """"
            setString5 = "SET v " & """" & u & """"
            setString6 = "SET p " & """" & x & """"
            setString7 = "SET vAd1 " & """" & uAd1 & """"
            setString8 = "SET vAd2 " & """" & uAd2 & """"
            setString9 = "SET pAd1 " & """" & xAd1 & """"
            setString10 = "SET pAd2 " & """" & xAd2 & """"
            setString11 = "SET vDesc " & """" & uDesc & """"
            setString12 = "SET pDesc " & """" & xDesc & """"

        Else
            setString1 = "SET v1 " & """" & x1 & """"
            setString2 = "SET v2 " & """" & x2 & """"
            setString3 = "SET p1 " & """" & u1 & """"
            setString4 = "SET p2 " & """" & u2 & """"
            setString5 = "SET v " & """" & x & """"
            setString6 = "SET p " & """" & u & """"
            setString7 = "SET vAd1 " & """" & xAd1 & """"
            setString8 = "SET vAd2 " & """" & xAd2 & """"
            setString9 = "SET pAd1 " & """" & uAd1 & """"
            setString10 = "SET pAd2 " & """" & uAd2 & """"
            setString11 = "SET vDesc " & """" & xDesc & """"
            setString12 = "SET pDesc " & """" & uDesc & """"
        End If

        addSetField(setString1)
        addSetField(setString2)
        addSetField(setString3)
        addSetField(setString4)
        addSetField(setString5)
        addSetField(setString6)
        addSetField(setString7)
        addSetField(setString8)
        addSetField(setString9)
        addSetField(setString10)
        addSetField(setString11)
        addSetField(setString12)

        ' set Exist variables based on if empty or not.
        ok = setExistVariables(u1Spouse, "u1Spouse")
        ok = setExistVariables(x2, "x2")
        ok = setExistVariables(u2, "u2")
        ok = setExistVariables(aMCo, "aMCo")
        ok = setExistVariables(nMCo, "nMCo")
        ok = setExistVariables(p1, "p2")
        RECalcFunction = NO_ERROR
        Exit Function

ErrorHandler:
        Dim errNum As Integer
        errNum = Err.Number

        Select Case errNum
            Case 0, 5941
                ' error zero is if we send it here, error 5941 is if "member of collection
                ' does not exist", most likely that the goMode bkm does not exist.
                msgBox("Error in exitDefaultOrHelpMode protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description & " possibly corrupted bookmarks")
                RECalcFunction = FUNCTION_ERROR
            Case Else
                msgBox("Error in exitDefaultOrHelpMode protected: " & "number: " _
                & Err.Number & "-message : " & Err.Description & " possibly corrupted bookmarks")
                RECalcFunction = FUNCTION_ERROR
        End Select
    End Function


    Function addSetField(ByVal setString As String)
        ' addSetField - Andrea 2008Jan
        ' DESCRIPTION: Used primarily by the calc Functions (similar to calc FILES R00-Calc)
        '   to read in all of the variables from the go file and make combinations as neccesary.
        '   This function writes to the go file a curly bracket field, with the text given
        '   in the input to the function, setString

        Selection.fields.Add(Range:=Selection.Range, Type:=wdFieldEmpty, Text:=setString)

    End Function


    Function setExistVariables(ByVal variable As String, ByVal varName As String)
        ' setExistVariables - Andrea 2008Jan
        ' DESCRIPTION: Used primarily by the calc functions.  This function checks to see
        '   if the variable in the input (name given by the input varName) is blank.  If it
        '   is, it writes a set field to the go file to set 'varName'&Exist to true, else
        '   sets it to false.  This varNameExist variable is then used throughout the
        '   programming to make decisions, ie if uSpouse is blank, dower, etc.

        Dim setString As String

        If (variable = "" Or variable = "     ") Then
            setString = "SET " & varName & "Exist " & """" & "f" & """"
        Else
            setString = "SET " & varName & "Exist " & """" & "t" & """"
        End If

        addSetField(setString)
        setExistVariables = "ok"
    End Function

    Function clearREFieldsAfterCalcFile() As String
        ' DESCRIPTION:  Checks if we need to clear the results from the formfields
        '   as listed below.  This step is neccesary, as during the merging process,
        '   if a user leaves these fields blank, the program fills them in with the
        '   a suggested value.  However, we do not want these values to hang around for
        '   the user to see - they didn't fill them in for a reason!  So this function
        '   checks to see if the program 'artificially wrote them in' by checking a bookmark
        '   set in the process of writing in values.  If the program wrote them in, then
        '   delete them.

        Dim blank, trueString As String
        On Error GoTo ErrorHandler

        ' set blank to five spaces, this is the value formfields are set to on blank.
        blank = "     "
        'trueString = "True"
        'ActiveDocument.Bookmarks("re2Delete").Range = trueString
        'ActiveDocument.Bookmarks.Exists ("re2Delete")

        If (ActiveDocument.Bookmarks.Exists("re2Delete")) Then
            ActiveDocument.FormFields("re2").result = blank
        End If
        If (ActiveDocument.Bookmarks.Exists("uAd1Delete")) Then
            ActiveDocument.FormFields("uAd1").result = blank
        End If
        If (ActiveDocument.Bookmarks.Exists("uAd2Delete")) Then
            ActiveDocument.FormFields("uAd2").result = blank
        End If
        If (ActiveDocument.Bookmarks.Exists("xAd1Delete")) Then
            ActiveDocument.FormFields("xAd1").result = blank
        End If
        If (ActiveDocument.Bookmarks.Exists("xAd2Delete")) Then
            ActiveDocument.FormFields("xAd2").result = blank
        End If
        If (ActiveDocument.Bookmarks.Exists("uDescDelete")) Then
            ActiveDocument.FormFields("uDesc").result = blank
        End If
        If (ActiveDocument.Bookmarks.Exists("xDescDelete")) Then
            ActiveDocument.FormFields("xDesc").result = blank
        End If
        clearREFieldsAfterCalcFile = NO_ERROR

        Exit Function
ErrorHandler:
        msgBox("Error in clearREFieldsAfterCalcFile: " & "number: " _
                & Err.Number & "-message : " & Err.Description)
        clearREFieldsAfterCalcFile = FUNCTION_ERROR

    End Function

End Module









