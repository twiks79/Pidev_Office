Sub Macro1()
    Dim currentRange As Range
    Dim foundTable As Boolean
    foundTable = False

    ' Set the initial search range from the current selection to the end of the document
    If Selection.Information(wdWithInTable) Then
        ' Start from the end of the current table
        Set currentRange = ActiveDocument.Range(Start:=Selection.Tables(1).Range.End, End:=ActiveDocument.Content.End)
    Else
        ' Start from the current selection position
        Set currentRange = ActiveDocument.Range(Start:=Selection.Range.End, End:=ActiveDocument.Content.End)
    End If

    ' Loop through all tables in the document
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        ' Check if the table start position is within the search range
        If tbl.Range.Start >= currentRange.Start Then
            ' Apply general styling to the found table
            With tbl
                .TopPadding = CentimetersToPoints(0.1)
                .BottomPadding = CentimetersToPoints(0.1)
                .LeftPadding = CentimetersToPoints(0.19)
                .RightPadding = CentimetersToPoints(0.1)
                .Spacing = 0
                .AllowPageBreaks = True
                .AllowAutoFit = True
            End With

            ' Loop through all cells in the table
            Dim cell As cell
            For Each cell In tbl.Range.Cells
                With cell.Range
                    ' Clear existing character formatting to ensure consistency
                    .Font.Size = 10 ' Set to your preferred size or retrieve from the style
                    .Font.Name = "Calibri" ' Set to your preferred font or retrieve from the style
                    
                    ' Check if the cell had bullet points originally
                    Dim hadBullet As Boolean
                    hadBullet = (.ListFormat.ListType <> WdListType.wdListNoNumbering)
                    
                    ' Apply "Body Text Table" style to the whole cell
                    .Style = ActiveDocument.Styles("Body Text Table")
                    
                    ' Reapply standard bullet points if the cell had them originally
                    If hadBullet Then
                        .ListFormat.ApplyBulletDefault
                        ' Set indentation properties for the bullet points
                        .ParagraphFormat.LeftIndent = CentimetersToPoints(0)
                        .ParagraphFormat.FirstLineIndent = CentimetersToPoints(-0.63)
                    End If
                End With
            Next cell
            
            ' Apply specific formatting to the first two cells in the header row
            Dim startRange As Range
            Dim endRange As Range
            Dim headerRange As Range
            Set startRange = tbl.cell(1, 1).Range
            Set endRange = tbl.cell(1, 2).Range
            endRange.MoveEnd wdCharacter, -1 ' Avoid including the cell's end-of-cell marker
            Set headerRange = ActiveDocument.Range(startRange.Start, endRange.End)
            With headerRange.Font
                .Bold = True
                .Color = wdColorWhite
            End With
            
            foundTable = True
            Exit For ' Remove this if you want to apply to all tables rather than stopping at the first found table
        End If
    Next tbl

    If Not foundTable Then
        MsgBox "No more tables found in the document."
    End If
End Sub
