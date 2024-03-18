Sub FindAndSaveCellTextsInNewTable()
    Dim doc As Document, newDoc As Document
    Dim tbl As Table, resultsTable As Table
    Dim cell As cell, nextCell As cell
    Dim rng As Range
    Dim userInput As String
    Dim cellIndex As Long
    Dim cellText As String
    Dim systemName As String
    
    ' Set the active document
    Set doc = ActiveDocument
    
    ' Ask the user for the text to compare
    userInput = InputBox("Enter the text to find in the tables:", "Text Input")
    
    ' Check if the user canceled the input box
    If userInput = "" Then Exit Sub
    
    ' Create a new document to display the results
    Set newDoc = Documents.Add
    ' Add a table to the new document
    Set resultsTable = newDoc.Tables.Add(Range:=newDoc.Range, NumRows:=1, NumColumns:=2)
    ' Set headers for the new table
    resultsTable.cell(1, 1).Range.Text = "System"
    resultsTable.cell(1, 2).Range.Text = userInput
    
    ' Loop through each table in the document
    For Each tbl In doc.Tables
        For Each cell In tbl.Range.Cells
            ' Compare cell text with user input, accounting for cell end characters
            If Trim(cell.Range.Text) = userInput & Chr(13) & Chr(7) Then
                ' Initialize cellIndex to avoid error if cell is the last one
                cellIndex = cell.Range.Start
                On Error Resume Next 'In case there is no cell to the right
                ' Try to get the next cell in sequence, could be below for vertically merged cells
                Set nextCell = doc.Range(Start:=cell.Range.End, End:=cell.Range.End + 1).Cells(1)
                If Err.Number = 0 And nextCell.Range.Start <> cellIndex Then 'Ensure it's a different cell
                    cellText = Replace(nextCell.Range.Text, Chr(13) & Chr(7), "")
                    If cellText <> "" Then
                    ' Find the nearest preceding heading level 3 for the system name
                    Set rng = cell.Range.Paragraphs(1).Range
                    rng.Collapse Direction:=wdCollapseStart
                    With rng.Find
                        .Style = doc.Styles("Heading 3")
                        .Forward = False ' Search backwards
                        .Execute
                        If .Found Then
                            systemName = rng.Text
                        Else
                            systemName = "N/A"
                        End If
                    End With

                        
                        ' Add a new row for the found text
                        With resultsTable
                            .Rows.Add
                            .cell(.Rows.Count, 1).Range.Text = systemName
                            .cell(.Rows.Count, 2).Range.Text = cellText
                        End With
                    End If
                End If
                On Error GoTo 0
            End If
        Next cell
    Next tbl
    
    ' Check if we added any rows beyond the header
    If resultsTable.Rows.Count = 1 Then
        MsgBox "No matching texts were found."
        newDoc.Close False
    Else
        MsgBox "Results are displayed in a new document with a table."
    End If
End Sub


