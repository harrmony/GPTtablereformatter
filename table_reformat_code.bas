Sub FormatTablesCustomStyle()
    Application.ScreenUpdating = False
    Dim tbl As Table
    Dim i As Long


    ' Step 1: Reformat Tables
    For i = 1 To ActiveDocument.Tables.Count
        Set tbl = ActiveDocument.Tables(i)


        ' Light formatting reset
        On Error Resume Next
        tbl.AutoFormat wdTableFormatNone
        tbl.AutoFitBehavior wdAutoFitFixed ' <-- Stop shrinking after AutoFormat
        On Error GoTo 0
        

        With tbl
            ' Set font and size for all text
            .Range.Font.Name = "Source Sans Pro"
            .Range.Font.Size = 11
            .Range.ParagraphFormat.SpaceAfter = 0
            .Range.ParagraphFormat.LeftIndent = CentimetersToPoints(0.23) ' <-- Add this line


            ' Set column widths
            If .Columns.Count >= 2 Then
                .Columns(1).PreferredWidthType = wdPreferredWidthPoints
                .Columns(1).PreferredWidth = CentimetersToPoints(8.5)
                .Columns(2).PreferredWidthType = wdPreferredWidthPoints
                .Columns(2).PreferredWidth = CentimetersToPoints(8.5)
            End If

            ' Set grey borders
            With .Borders
                .Enable = True
                .InsideLineStyle = wdLineStyleSingle
                .OutsideLineStyle = wdLineStyleSingle
                .InsideColor = RGB(150, 150, 150) ' Grey
                .OutsideColor = RGB(150, 150, 150)
            End With

            ' Style top row (header)
            With .Rows(1).Range
                .Font.Bold = True
                .Font.Color = RGB(21, 96, 130) ' RL Dark blue is RGB(91, 155, 213)
            End With

            ' Style remaining rows
            Dim r As Long
            For r = 2 To .Rows.Count
                With .Cell(r, 1).Range
                    .Font.Bold = False
                    .Font.Color = RGB(102, 102, 102) ' Grey
                End With
                If .Columns.Count >= 2 Then
                    With .Cell(r, 2).Range
                        .Font.Bold = True
                        .Font.Color = RGB(21, 96, 130) ' RL Dark blue is RGB(91, 155, 213) -- Oral-B RGB(21, 96, 130)
                    End With
                End If
            Next r
        End With
    Next i

' Step 2: Recolour all [bracketed] text in the document
Dim findRng As Range
Set findRng = ActiveDocument.Content

With findRng.Find
    .ClearFormatting
    .Text = "\[*\]"
    .MatchWildcards = True
    .Forward = True
    .Wrap = wdFindStop
    Do While .Execute
        ' Apply color to found text (includes the brackets)
        findRng.Font.Color = RGB(21, 96, 130) ' RL Dark blue is RGB(91, 155, 213)
        findRng.Font.Bold = True
        findRng.Collapse wdCollapseEnd
    Loop
End With

    Application.ScreenUpdating = True
    MsgBox "All tables formatted and bracketed text recoloured."

End Sub

