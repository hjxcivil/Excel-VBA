# Using the Iif Function in VBA

[TOC]

## The If Statement

Sub *BasicIfStatement*()

    Dim Wins As Integer
    Dim Success As String
    
    Range("A6").Select
    
    Wins = ActiveCell.Offset(0, 10).Value
    
    If Wins >= 1 Then Success = "Winner" Else Success = "Loser"
    
    
    Debug.Print Wins, Success

End Sub

## The Iif Statement

`Success = IIf(Wins >= 1, "Winner", "Loser")`

## Writing Nested Ifs and Nested Iifs

- Creating a Nested If Statement

  ```
  If Wins >= 1 Then
  	Success = "Winner"
  Else
      If Noms >= 1 Then
          Success = "Loser"
      Else
          Success = "Nobody"
      End If
  End If
  ```

  

- Creating a Nested Iif Function

  `Success = IIf(Wins >= 1, "Winner", IIf(Noms >= 1, "Loser", "Nobody"))`

## A Practical Example Using Iif

Sub *ListFilmsBySuccess*()

    Dim Wins As Integer, Noms As Integer
    Dim Success As String
    Dim r As Range
    Dim SheetNames()
    Dim v As Variant
    Dim Cols As Integer
    
    Cols = Sheet1.Range("A1").CurrentRegion.Columns.Count
    
    SheetNames = Array("Winner", "Loser", "Nobody")
    
    For Each v In SheetNames
        ClearSheet v
    Next v
    
    For Each r In Sheet1.Range("A2", Sheet1.Range("A1").End(xlDown))
    
        Wins = r.Offset(0, 10).Value
        Noms = r.Offset(0, 9).Value
        
        Success = IIf(Wins >= 1, "Winner", IIf(Noms >= 1, "Loser", "Nobody"))
    
        r.Resize(1, Cols).Copy Worksheets(Success).Range("A1048576").End(xlUp).Offset(1, 0)
        
    Next r
    
    For Each v In SheetNames
        Worksheets(v).Range("A1").CurrentRegion.EntireColumn.AutoFit
    Next v

End Sub

Sub *ClearSheet*(ByVal *SheetName* As String)

    On Error GoTo CreateSheet
    Worksheets(SheetName).Cells.Clear
    On Error GoTo 0
    
    Sheet1.Range("A1", Sheet1.Range("A1").End(xlToRight)).Copy Worksheets(SheetName).Range("A1")
    
    Exit Sub

```
CreateSheet:
    Worksheets.Add.Name = SheetName
    Resume Next
```


End Sub