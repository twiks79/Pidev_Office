Sub Macro1()
    Dim targetRange As Range
    Dim tbl As Table
    
    ' Determine if there's a selection that includes more than just a cursor position (i.e., selection length > 1)
    If Selection.Range.Start <> Selection.Range.End Then
        ' Use the selected range if there's a selection
        Set targetRange = Selection.Range
    Else
        ' Use the entire document if there's no specific selection
        Set targetRange = ActiveDocument.Content
    End If
    
    ' Variable to track if at least one table is formatted
    Dim foundTable As Boolean
    foundTable = False
    
    ' Loop through all tables in the target range
    For Each tbl In targetRange.Tables
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
                    .ParagraphFormat.LeftIndent = CentimetersToPoints(0.63) ' Moves start of text to the right
                    .ParagraphFormat.FirstLineIndent = CentimetersToPoints(-0.63) ' Aligns bullets to the left, with text indented from bullets

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
    Next tbl

    If Not foundTable Then
        MsgBox "No tables found in the target range."
    End If
End Sub
