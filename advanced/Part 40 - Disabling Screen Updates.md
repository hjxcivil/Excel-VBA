# Disabling Screen Updates

[TOC]

## Turning Off Screen Updates

Sub *ColourInABunchOfCells*()

    Dim ws As Worksheet
    Dim r As Range
    
    Application.ScreenUpdating = False ' 2.00s vs 37.99s
    
    For Each ws In Worksheets
    
        ws.Select
        
        For Each r In Range("A1:z50")
                
            r.Select
            r.Interior.Color = rgbRed ' rgbGreen ' rgbRed
            
        Next r
        
    Next ws

End Sub

## Creating a Simple Timer

Sub *ColourInABunchOfCells*()

    Dim ws As Worksheet
    Dim r As Range
    Dim StartTime As Date, EndTime As Date
    Dim TimeTaken As Double
    
    StartTime = Time
    
    Application.ScreenUpdating = False ' 3.00s vs 41.99s
    
    For Each ws In Worksheets
    
        ws.Select
        
        For Each r In Range("A1:z50")
                
            r.Select
            r.Interior.Color = rgbBlue ' rgbGreen ' rgbRed
            
        Next r
        
    Next ws
    
    EndTime = Time
    
    TimeTaken = (EndTime - StartTime) * 24 * 60 * 60
    
    MsgBox TimeTaken

End Sub

## A More Complex Example

Sub *SeparateFilmsByGenre*()

    Dim Genre As String
    Dim StartTime As Date, EndTime As Date
    Dim TimeTaken As Double
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False '  97.99s vs 309s
    StartTime = Time
    
    wsMovies.Select
    Range("A2").Select
    
    Do Until ActiveCell.Value = ""
        
        Genre = ActiveCell.Offset(0, 7).Value
        
        If Not SheetExists(Genre) Then
            Worksheets.Add After:=Sheets(Sheets.Count)
            ActiveSheet.Name = Genre
            wsMovies.Range("A1").EntireRow.Copy ActiveCell
            Range("A2").Select
        End If
        
        wsMovies.Select
        ActiveCell.EntireRow.Copy
        Worksheets(Genre).Select
        ActiveCell.PasteSpecial
        ActiveCell.Offset(1, 0).Select
        
        wsMovies.Select
        ActiveCell.Offset(1, 0).Select
    Loop
    
    Application.CutCopyMode = False
    
    For Each ws In Worksheets
        ws.Select
        Range("A1").Select
        ActiveCell.CurrentRegion.EntireColumn.AutoFit
    Next ws
    
    wsMovies.Select
    
    EndTime = Time
    TimeTaken = (EndTime - StartTime) * 24 * 60 * 60
    
    MsgBox TimeTaken & " seconds"

End Sub

Function *SheetExists*(SheetName As String) As Boolean
    
    On Error GoTo NoSheet
    Sheets(SheetName).Select
    SheetExists = True
    
    Exit Function
    
    NoSheet:
        SheetExists = False

End Function

Sub *DeleteAllButMovies*()

    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In Worksheets
        If Not ws Is wsMovies Then ws.Delete
    Next ws

End Sub

Sub *SeparateFilmsAvoidingSelectingThings*() 'equal to off ScreenUpdating

    Dim Genre As String
    Dim StartTime As Date, EndTime As Date
    Dim TimeTaken As Double
    Dim ws As Worksheet
    Dim r As Range, rs As Range
    
    StartTime = Time
    
    Set rs = _
        wsMovies.Range("A2", wsMovies.Range("A1").End(xlDown))
    
    For Each r In rs
        
        Genre = r.Offset(0, 7).Value
        
        If Not SheetExists(Genre) Then
            Worksheets.Add After:=Sheets(Sheets.Count)
            ActiveSheet.Name = Genre
            wsMovies.Range("A1").EntireRow.Copy ActiveCell
        End If
        
        r.EntireRow.Copy _
            Worksheets(Genre).Range("A1048576").End(xlUp).Offset(1, 0)
        
    Next r
    
    Application.CutCopyMode = False
    
    For Each ws In Worksheets
        ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Next ws
        
    EndTime = Time
    TimeTaken = (EndTime - StartTime) * 24 * 60 * 60
    
    MsgBox TimeTaken & " seconds"

End Sub