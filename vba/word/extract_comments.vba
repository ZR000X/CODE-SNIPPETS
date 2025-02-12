Sub ExportCommentsToExcel()
    Dim xlApp As Object
    Dim xlWB As Object
    Dim i As Integer
    Dim commentRange As Range
    Dim sectionTitle As String
    Dim para As Paragraph
    Dim paraStyle As String
    Dim paraLevel As Integer
    Dim paraText As String
    Dim referenceText As String
    
    ' Create a new instance of Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Add
    
    ' Add headers to the Excel sheet
    With xlWB.Worksheets(1)
        .Cells(1, 1).Value = "Author Name"
        .Cells(1, 2).Value = "Section"
        .Cells(1, 3).Value = "Reference Text"
        .Cells(1, 4).Value = "Comment Text"
    End With
    
    ' Loop through each comment in the Word document
    For i = 1 To ActiveDocument.Comments.Count
        Set commentRange = ActiveDocument.Comments(i).Scope
        sectionTitle = "No Section Found"
        referenceText = Trim(commentRange.Text) ' Get the reference text
        
        ' Traverse backwards through paragraphs to find the nearest heading
        For Each para In ActiveDocument.Paragraphs
            If para.Range.Start <= commentRange.Start Then
                paraStyle = para.Style
                If InStr(1, paraStyle, "Heading") > 0 Then
                    paraLevel = Val(Mid(paraStyle, 8))
                    If paraLevel > 0 Then
                        paraText = Trim(para.Range.Text)
                        sectionTitle = paraText
                        ' Continue searching for a more specific (lower level) heading
                    End If
                End If
            End If
        Next para
        
        ' Write the comment details to the Excel sheet
        With xlWB.Worksheets(1)
            .Cells(i + 1, 1).Value = ActiveDocument.Comments(i).Author
            .Cells(i + 1, 2).Value = sectionTitle
            .Cells(i + 1, 3).Value = referenceText
            .Cells(i + 1, 4).Value = ActiveDocument.Comments(i).Range.Text
        End With
    Next i
    
    ' Release Excel objects
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub
