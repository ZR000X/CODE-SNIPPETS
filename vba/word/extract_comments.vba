Sub ExportCommentsToExcel()
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim i As Long                    ' Changed to Long for larger documents
    Dim commentRange As Range
    Dim sectionTitle As String
    Dim para As Paragraph
    Dim paraStyle As String
    Dim paraLevel As Integer
    Dim paraText As String
    Dim headingCache As Collection   ' New: Cache for headings
    Dim statusMsg As String
    Dim startTime As Double
    
    ' Initialize progress tracking
    startTime = Timer
    Application.ScreenUpdating = False
    
    ' Create heading cache for better performance
    Set headingCache = New Collection
    Call CacheHeadings(headingCache)
    
    ' Create and configure Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False  ' Hide until ready
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Worksheets(1)
    
    ' Add headers and format them
    With xlSheet
        .Cells(1, 1).Value = "Author Name"
        .Cells(1, 2).Value = "Section"
        .Cells(1, 3).Value = "Comment Text"
        .Cells(1, 4).Value = "Page Number"    ' New column
        .Cells(1, 5).Value = "Date/Time"      ' New column
        
        ' Format headers
        With .Range("A1:E1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
    End With
    
    ' Show progress bar
    Application.StatusBar = "Processing comments... Please wait."
    
    ' Loop through comments
    For i = 1 To ActiveDocument.Comments.Count
        Set commentRange = ActiveDocument.Comments(i).Scope
        sectionTitle = FindNearestHeading(commentRange.Start, headingCache)
        
        ' Write to Excel with additional information
        With xlSheet
            .Cells(i + 1, 1).Value = ActiveDocument.Comments(i).Author
            .Cells(i + 1, 2).Value = sectionTitle
            .Cells(i + 1, 3).Value = ActiveDocument.Comments(i).Range.Text
            .Cells(i + 1, 4).Value = commentRange.Information(wdActiveEndPageNumber)
            .Cells(i + 1, 5).Value = Format(ActiveDocument.Comments(i).Date, "yyyy-mm-dd hh:mm")
        End With
        
        ' Update progress every 10 comments
        If i Mod 10 = 0 Then
            statusMsg = "Processing comment " & i & " of " & ActiveDocument.Comments.Count
            Application.StatusBar = statusMsg & " (" & Format((i / ActiveDocument.Comments.Count) * 100, "0") & "%)"
        End If
    Next i
    
    ' Format Excel sheet
    With xlSheet
        .Columns("A:E").AutoFit
        .Range("A1:E" & ActiveDocument.Comments.Count + 1).Borders.LineStyle = 1 'xlContinuous
        
        ' Fix sorting with proper Excel constants
        .Range("A2:E" & ActiveDocument.Comments.Count + 1).Sort _
            Key1:=.Range("B2"), _
            Order1:=1, _  'xlAscending = 1
            Header:=1     'xlYes = 1
    End With
    
    ' Show Excel and clean up
    xlApp.Visible = True
    xlSheet.Activate
    xlSheet.Range("A1").Select
    
CleanUp:
    ' Restore Word settings
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Release objects
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Set headingCache = Nothing
    
    ' Show completion message
    MsgBox "Export completed successfully!" & vbNewLine & _
           "Processed " & ActiveDocument.Comments.Count & " comments in " & _
           Format(Timer - startTime, "0.0") & " seconds.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' New helper function to cache headings for better performance
Private Sub CacheHeadings(ByRef headingCache As Collection)
    Dim para As Paragraph
    Dim paraStyle As String
    Dim headingInfo As String
    
    For Each para In ActiveDocument.Paragraphs
        paraStyle = para.Style
        If InStr(1, paraStyle, "Heading") > 0 Then
            headingInfo = para.Range.Start & "|" & Trim(para.Range.Text)
            headingCache.Add headingInfo, CStr(para.Range.Start)
        End If
    Next para
End Sub

' New helper function to find nearest heading using cache
Private Function FindNearestHeading(ByVal position As Long, ByRef headingCache As Collection) As String
    Dim i As Long
    Dim nearestHeading As String
    Dim headingInfo() As String
    
    nearestHeading = "No Section Found"
    
    ' Search through cached headings
    For i = headingCache.Count To 1 Step -1
        headingInfo = Split(headingCache.Item(i), "|")
        If CLng(headingInfo(0)) <= position Then
            nearestHeading = headingInfo(1)
            Exit For
        End If
    Next i
    
    FindNearestHeading = nearestHeading
End Function
