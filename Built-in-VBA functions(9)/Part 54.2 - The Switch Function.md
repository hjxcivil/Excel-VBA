# Using the Switch Function in VBA

[TOC]

## A Reminder of the Select Case Statement

Sub *BasicSelectCase*()

    Dim RunTime As Integer
    Dim Length As String
    
    Range("A4").Select
    RunTime = ActiveCell.Offset(0, 3).Value
    
    Select Case RunTime
        Case Is <= 90
            Length = "Short"
        Case Is <= 120
            Length = "Medium"
        Case Is <= 150
            Length = "Long"
        Case Is <= 180
            Length = "Epic"
        Case Else
            Length = "Snooze fest"
    End Select
    
    Debug.Print RunTime, Length

End Sub

## Writing a Basic Switch Function

Sub *BasicSwitch*()

    Dim RunTime As Integer
    Dim Length As String
    
    Range("A4").Select
    RunTime = ActiveCell.Offset(0, 3).Value
    
    Length = Switch( _
         RunTime <= 90, "Short", _
        RunTime <= 120, "Medium", _
        RunTime <= 150, "Long", _
        RunTime <= 180, "Epic", _
        True, "Snooze fest")
    
    Debug.Print RunTime, Length

End Sub

## Using Switch in a User-Defined Function

`Length = FilmLength(RunTime)`

Function *FilmLength*(*RunTime* As Integer) As String

    FilmLength = Switch( _
        RunTime <= 90, "Short", _
        RunTime <= 120, "Medium", _
        RunTime <= 150, "Long", _
        RunTime <= 180, "Epic", _
        True, "Snooze fest")

End Function

## A Practical Example

Sub *ListFilmsByLength*()
    
    Dim r As Range
    Dim Length As String
    Dim SheetNames()
    Dim v As Variant
    Dim Cols As Integer
    
    Cols = Sheet1.Range("A1").CurrentRegion.Columns.Count
    
    SheetNames = Array("Short", "Medium", "Long", "Epic", "Snooze fest")
    
    For Each v In SheetNames
        CreateLengthSheet v
    Next v
    
    For Each r In Sheet1.Range("A2", Sheet1.Range("A1").End(xlDown))
        Length = FilmLength(r.Offset(0, 3).Value)
        r.Resize(1, Cols).Copy Worksheets(Length).Range("A1048576").End(xlUp).Offset(1, 0)
    Next r
    
    For Each v In SheetNames
        Worksheets(v).Range("A1").CurrentRegion.EntireColumn.AutoFit
    Next v

End Sub

Sub *CreateLengthSheet*(ByVal *SheetName* As String)
    
    On Error GoTo createsheet
    Worksheets(SheetName).Cells.Clear
    On Error GoTo 0
    
    Sheet1.Range("A1", Sheet1.Range("A1").End(xlToRight)).Copy Worksheets(SheetName).Range("A1")
    
    Exit Sub

```
createsheet:
    Worksheets.Add.Name = SheetName
    Resume Next
```


​    
End Sub