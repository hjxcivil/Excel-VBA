# VBA Arrays

[TOC]

## What is a Array?

## Declaring a Fixed-Size Array

​	`Dim TopThreeFilms(2) As String`

- Using the Option Base Statement

  `Option Base 0`
  
- Declaring the Lower and Upper Bounds

  `Dim TopThreeFilms(1 to 3) As String`

## Writing to and Reading from an Array

Sub *FixedSizeArray*()

    Dim TopThreeFilms(1 To 3) As String
    
    TopThreeFilms(1) = Range("B3").Value
    TopThreeFilms(2) = Range("B4").Value
    TopThreeFilms(3) = Range("B5").Value
    
    Worksheets.Add
    
    Range("A1").Value = TopThreeFilms(3)
    Range("A2").Value = TopThreeFilms(2)
    Range("A3").Value = TopThreeFilms(1)
    
    Erase TopThreeFilms

End Sub

## Looping Over an Array

Sub *LoopOverArray*()

    Dim TopTenFilms(1 To 13) As String
    Dim Counter As Long
    
    Sheet1.Activate
    
    For Counter = LBound(TopTenFilms) To UBound(TopTenFilms)
        TopTenFilms(Counter) = Range("B2").Offset(Counter, 0).Value
    Next Counter
    
    Worksheets.Add
    
    For Counter = UBound(TopTenFilms) To LBound(TopTenFilms) Step -1
        ActiveCell.Value = TopTenFilms(Counter)
        ActiveCell.Offset(1, 0).Select
    Next Counter
    
    Erase TopTenFilms

End Sub

## Erasing Arrays

​	`Erase TopTenFilms`

## Multi-Dimension Arrays

Sub *MultiDimensionArray*()

    Dim TopFilms(0 To 9, 0 To 4) As Variant
    Dim Dimension1 As Long, Dimension2 As Long
    
    For Dimension1 = LBound(TopFilms, 1) To UBound(TopFilms, 1)
        For Dimension2 = LBound(TopFilms, 2) To UBound(TopFilms, 2)
            TopFilms(Dimension1, Dimension2) = Range("A3").Offset(Dimension1, Dimension2).Value
        Next Dimension2
    Next Dimension1
    
    Worksheets.Add
    
    For Dimension1 = LBound(TopFilms, 1) To UBound(TopFilms, 1)
        For Dimension2 = LBound(TopFilms, 2) To UBound(TopFilms, 2)
             ActiveCell.Offset(Dimension1, Dimension2).Value = TopFilms(Dimension1, Dimension2)
        Next Dimension2
    Next Dimension1
    
    Erase TopFilms

End Sub

## Dynamic Arrays

Sub *DynamicMultiDimensionArray*

    Dim TopFilms() As Variant
    Dim Dimension1 As Long, Dimension2 As Long
    
    Sheet1.Activate
    
    Dimension1 = Range("A3", Range("A2").End(xlDown)).Cells.Count - 1
    Dimension2 = Range("A2", Range("A2").End(xlToRight)).Cells.Count - 1
    
    ReDim TopFilms(0 To Dimension1, 0 To Dimension2)
    ...
End Sub

- Writing a Range to a Dynamic Array '1 Base

  Sub *QuickDynamicMultiDimensionArray*()

  ```
  Dim TopFilms() As Variant
  
  Sheet1.Activate
  
  TopFilms = Range("A3", Range("A2").End(xlDown).End(xlToRight))
  
  ...
  
  Erase TopFilms
  ```

  End Sub

- Erasing Dynamic Arrays 'Clear not empty 

- Writing a Dynamic Array to a Range

  ```
  Worksheets.Add
  
  Range(ActiveCell, ActiveCell.Offset(UBound(TopFilms, 1) - 1, UBound(TopFilms, 2) - 1)).Value = TopFilms
  ```

## Performing Calculations in Arrays

Sub *CalculateWithArray*()

    Dim FilmLength() As Variant
    Dim Answers() As Variant
    Dim Dimension1 As Long, Counter As Long
    
    Sheet1.Activate
    
    FilmLength = Range("D3", Range("D2").End(xlDown))
    
    Dimension1 = UBound(FilmLength, 1)
    
    ReDim Answers(1 To Dimension1, 1 To 2)
    
    For Counter = 1 To Dimension1
        Answers(Counter, 1) = Int(FilmLength(Counter, 1) / 60)
        Answers(Counter, 2) = FilmLength(Counter, 1) Mod 60
    Next Counter
    
    Range("F3", Range("F3").Offset(Dimension1 - 1, 1)).Value = Answers
    
    Erase FilmLength
    Erase Answers

End Sub

## Resizing Arrays Dynamically

Sub *ResizeDynamicArray*()
    Dim ActionFilms() As Variant
    Dim r As Range
    Dim ActionCounter As Long, LoopCounter As Long
        
    Sheet1.Activate
    
    For Each r In Range("A3", Range("A2").End(xlDown))
        If LCase(r.Offset(0, 4).Value) = "action" Then
            
            ActionCounter = ActionCounter + 1
            
            ReDim Preserve ActionFilms(1 To 5, 1 To ActionCounter)
            
            For LoopCounter = 1 To 5
                ActionFilms(LoopCounter, ActionCounter) = r.Offset(0, LoopCounter - 1).Value
            Next LoopCounter
            
        End If
    Next r
    
    Worksheets.Add
    
    Range(ActiveCell, ActiveCell.Offset(UBound(ActionFilms, 2) - 1, 4)).Value = Application.Transpose(ActionFilms)

End Sub