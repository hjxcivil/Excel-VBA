# Using Object Variables in VBA

[TOC]

## Declaring Object Variables

- A Recap Of Basic Variables

  Sub *RecapOfBasicVariables*()

      Dim MyStringVariable As String
      Dim MyIntegerVariable As Integer
      Dim MyDateVariable As Date
      
      MyStringVariable = "Some text"
      MyIntegerVariable = 1
      MyDateVariable = Now
      
      MsgBox "String variable contains: " & MyStringVariable
      MsgBox "Integer variable contains: " & MyIntegerVariable
      MsgBox "Date variable contains: " & MyDateVariable

  End Sub

- Declaring and Setting Object Variables

  Sub *StoreRangeOfCells*() 

      Dim FilmNameCells As Range
      
      Set FilmNameCells = Range("B3", Range("B3").End(xlDown)) 'Advantage 1
      
      Sheet2.Activate 'Advantage2
      
      FilmNameCells.Font.Color = rgbRed
      FilmNameCells.Font.Italic = False

  End Sub

- Storing New Objects in Variables

  Sub *ReferencingAWorksheetInAVariable*()

      Dim MyNewSheet As Worksheet
      
      Set MyNewSheet = Worksheets.Add
      
      Sheet1.Activate
      Range("A1").CurrentRegion.Copy
      
      MyNewSheet.Activate
      ActiveCell.PasteSpecial

  End Sub

## Creating and Referencing Objects

Sub *OtherEgs*()

    Dim MyNewBook As Workbook
    
    Set MyNewBook = Workbooks.Add("Top Movies 2012.xltm")
    
    Dim MyNewChart As Chart
    
    Set MyNewChart = Charts.Add

End Sub

## Finding and Referencing a Range

Sub *FindingARange*()

    Dim FilmToFind As String
    Dim FilmCell As Range
    
    FilmToFind = InputBox("Type in a film name")
    
    Set FilmCell = _
        Range("B3", Range("B3").End(xlDown)).Find(FilmToFind)
    
    If FilmCell Is Nothing Then
        MsgBox FilmToFind & " was not found"
    Else
        MsgBox FilmCell.Value & " was found in " & FilmCell.Address
    End If

End Sub