Module NewMacros

Option Explicit On
    ' MODULE DESCRIPTION: New Macros.  When the user creates/records a new macro
    '   through the word interface (tools->record/view macro) it goes to this file.
    '   As most often recordings are done to just see what the code for an action is,
    '   and are disposable, this module should be cleaned out from time to time by
    '   deleting old functions, and classifying others.

    Sub aaaaaopenDaytimer()
        '
        ' aaaaaopenDaytimer Macro
        ' Macro created 23/10/2010 by Grant
        '

    End Sub
    Sub xxRecording()
        '
        ' xxRecording Macro
        ' Macro recorded 23/10/2010 by Grant
        '
        If ActiveWindow.View.SplitSpecial = wdPaneNone Then
            ActiveWindow.ActivePane.View.Type = wdNormalView
        Else
            ActiveWindow.View.Type = wdNormalView
        End If
    End Sub
    Sub xxxRecordingAgain()
        '
        ' xxxRecordingAgain Macro
        ' Macro recorded 23/10/2010 by Grant
        '
        Application.Run MacroName:="Normal.General.toggleCodes"
        ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View. _
            ShowAll
    End Sub
    Sub Macro1()
        '
        ' Macro1 Macro
        ' Macro recorded 23/10/2010 by Grant
        '
    End Sub
    Sub aaaChangeAllTimesNewRoman12()
        '
        ' ChangeAllTimesNewRoman12 Macro
        ' Macro recorded 23/10/2010 by Grant
        '
        Selection.WholeStory()
        Selection.Font.name = "Times New Roman"
        Selection.Font.Size = 12
    End Sub
    Sub makeAllSingleSpace()
        '
        ' makeAllSingleSpace Macro
        ' Macro recorded 23/10/2010 by Grant
        '
        Selection.WholeStory()
        With Selection.ParagraphFormat
            .SpaceBeforeAuto = False
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
        End With
    End Sub
    Sub insertSectionBreakEndAndStar()
        '
        ' insertSectionBreakEndAndStar Macro
        ' Macro recorded 23/10/2010 by Grant
        '
        Selection.EndKey Unit:=wdStory
        Selection.InsertBreak Type:=wdSectionBreakNextPage
        Selection.TypeText Text:="*"
    End Sub
    Sub AssignBodyBookmark()
        '
        ' AssignBodyBookmark Macro
        ' Macro recorded 25/10/2010 by Grant
        '
        Selection.HomeKey Unit:=wdStory
        Selection.EndKey(Unit:=wdStory, Extend:=wdExtend)
        Selection.MoveUp(Unit:=wdLine, count:=1, Extend:=wdExtend)
        With ActiveDocument.Bookmarks
            .Add(Range:=Selection.Range, name:="body")
            .DefaultSorting = wdSortByName
            .ShowHidden = False
        End With
    End Sub
    Sub InsertHeaderStuffRecording()
        '
        ' InsertHeaderStuffRecording Macro
        ' Macro recorded 25/10/2010 by Grant
        '
        If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            ActiveWindow.Panes(2).Close()
        End If
        If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
            ActivePane.View.Type = wdOutlineView Then
            ActiveWindow.ActivePane.View.Type = wdPrintView
        End If
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            ActiveWindow.Panes(2).Close()
        End If
        If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
            ActivePane.View.Type = wdOutlineView Then
            ActiveWindow.ActivePane.View.Type = wdPrintView
        End If
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        ActiveDocument.Save()
    End Sub
    Sub SetFirstPageDifferentHeader()
        '
        ' Macro2 Macro
        ' Macro recorded 25/10/2010 by Grant
        '
        With Selection.PageSetup
            .DifferentFirstPageHeaderFooter = True
        End With
    End Sub
    Sub Macro2()
        '
        ' Macro2 Macro
        ' Macro recorded 29/10/2010 by Grant
        '
        Selection.MoveDown(Unit:=wdLine, count:=1)
        Selection.MoveLeft(Unit:=wdCharacter, count:=4)
        Selection.MoveUp(Unit:=wdLine, count:=1)
    End Sub
    Sub SelectAllTimesNewRoman10Save()
        '
        ' SelectAllTimesNewRoman12Save Macro
        ' Macro recorded 30/10/2010 by Grant
        '
        Selection.WholeStory()
        ActiveWindow.ActivePane.VerticalPercentScrolled = 0
        Selection.Font.name = "Times New Roman"
        Selection.Font.Size = 10
        ActiveDocument.Save()
    End Sub
    Sub xxx()
        '
        ' xxx Macro
        ' Macro recorded 10/07/2013 by Grant
        '
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    End Sub


End Module
