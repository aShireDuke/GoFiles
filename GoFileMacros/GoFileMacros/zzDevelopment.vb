Module zzDevelopment

Option Explicit On

    '***********************************************************
    ' MODULE DESCRIPTION:  FormReferenceCompanyNames.  This module contains all of the routines that relate to
    '   the 'databasing' system of realty, mortgage, and law office names.
    '***********************************************************
    Sub openReferenceDoc()

        Dim goDoc, referenceDoc As Document
        goDoc = ActiveDocument

        Documents.Open("zReferenceDoc.doc")
        referenceDoc = ActiveDocument
        'goDoc.Activate
        goFileName = goDoc.name

        'now wait for user to highlight selection
        'or can close after time

    End Sub


    Function readHighlightedReferenceSection()

        'on click of button (one located in reference doc) can grab whole
        'highlighted section of table & read & insert into goDoc.

        'get values into variables

        'activate goFileName (as have link to global)

        'change reCo, address, etc result to variables just grabbed

    End Function

    Function enableFieldsOnCheck(ByVal fields, ByVal checkName As String) As String
        ' enableFieldsOnCheck - Andrea December 2007
        ' DESCRIPTION:  This function gives the user the ability to lock a whole group
        '   of fields to not be able to edit until a controlling check Box is activated.
        '   The user inputs an array of strings (fields) containing all of the names the
        '   that the button will enable/disable.  Also inputted to this function is the
        '   name of the checkbox that will activate the input settings of the field.

        Dim checkValue As Boolean
        Dim currentField As String
        Dim J, fieldSize As Integer
        On Error GoTo ErrorHandler
        ' read current value of the checkBox, and enable/disable ability of
        ' user to input data into the specified formfields
        checkValue = ActiveDocument.FormFields(checkName).CheckBox.value
        fieldSize = UBound(fields)
        If (checkValue = True) Then
            ' If the the checkbox is checked, allow modification of formfields
            ' Read in name of the fields array bookmark and enable
            For J = 0 To fieldSize
                ActiveDocument.FormFields(fields(J)).Enabled = True
            Next J
        Else
            ' If the checkbox is unchecked, keep disabled
            For J = 0 To fieldSize
                ActiveDocument.FormFields(fields(J)).Enabled = False
            Next J
        End If

    End Function

    Sub enableOrDisableUSpouseOnExit()

        Dim dowerStatus As String

        ' Write the information from the field you were just in to the bookmark
        setDefaultFormField()

        'read from the bookmark & make decision based on it.

        dowerStatus = ActiveDocument.FormFields("dowerStatus").result
        Select Case dowerStatus
            Case "Married - Joint Tenants"
                'then want to enable uSpouse fill in
                ActiveDocument.FormFields("u1Spouse").Enabled = True

            Case "Not Married"
                ActiveDocument.FormFields("u1Spouse").Enabled = False
                ActiveDocument.FormFields("u1Spouse").result = "NotMarried"

            Case "Married - Remove Title"
                ActiveDocument.FormFields("u1Spouse").Enabled = False
                ActiveDocument.FormFields("u1Spouse").result = "NotMarried"
            Case Else
        End Select

    End Sub


    Sub enableUFirmFieldsOnCheck()
        ' enableUFirmFieldsOnCheck - Andrea Dec 2007
        ' DESCRIPTION: This function enables/disables a set of fields using the
        '   function enableFieldsOnCheck found in FormDocumentControls module

        Dim fields(14) As String
        Dim status As Integer

        fields(0) = "uFirm"
        fields(1) = "uLawyer"
        fields(2) = "uFPhNo"
        fields(3) = "uFAd1"
        fields(4) = "uFAd2"
        fields(5) = "uFCode"
        fields(6) = "uFGSTNo"
        fields(7) = "witness1"
        fields(8) = "wit1Title"
        fields(9) = "wit1Ad1"
        fields(10) = "wit1Ad2"
        fields(11) = "witness2"
        fields(12) = "wit2Title"
        fields(13) = "wit2Ad1"
        fields(14) = "wit2Ad2"

        status = enableFieldsOnCheck(fields, "OurFirmCheckBox")
    End Sub

    Sub RemoveAllHeadersAndFooters()
        ' RemoveAllHeadersAndFooters - Andrea Oct 2010
        ' DESCRIPTION: Deletes all headers from :http://word.tips.net/Pages/T001777_Deleting_All_Headers_and_Footers.html
        Dim oSec As Section
        Dim oHead As HeaderFooter
        Dim oFoot As HeaderFooter

        For Each oSec In ActiveDocument.Sections
            For Each oHead In oSec.Headers
                If oHead.Exists Then oHead.Range.Delete()
            Next oHead

            For Each oFoot In oSec.Footers
                If oFoot.Exists Then oFoot.Range.Delete()
            Next oFoot
        Next oSec
    End Sub



End Module
