# Using the Application.InputBox Method

[TOC]

## Limitations of the Generic InputBox Function

- No way to ctrl the data set returned and that can lead to runtime errors

  `FilmDate = InputBox("Enter a date") 'fefefefef`'

- modal - dialog: no interact

## Using the Application.InputBox

`FilmName = Application.InputBox("Enter a film name")`

## Setting the Return Type

| Value | Description                             |
| :---- | :-------------------------------------- |
| 0     | A formula                               |
| 1     | A number                                |
| 2     | Text (a string)                         |
| 4     | A logical value (**True** or **False**) |
| 8     | A cell reference, as a **Range** object |
| 16    | An error value, such as #N/A            |
| 64    | An array of values                      |

## Validating User Input

- Returning a Number

  `FilmLength = Application.InputBox(Prompt:="Enter the length", Type:=1)`

- Dealing with Dates

  `FilmDate = Application.InputBox(Prompt:="Enter a date dd/mm/yyyy", Type:=1)`

- Returning a formula

  `Myformula = Application.InputBox(Prompt:="Enter a formula", Type:=0, Default:="=SUM(")`

- Returning an Array

- Returning a Range

  `Set FormulaCell = Application.InputBox(Prompt:="Choose formula cell", Type:=8)`

  `FormulaCell.FormulaLocal = Myformula`

  - Multiple Cells

    ```
    Set CopyRange = Application.InputBox(Prompt:="Choose cells to copy", Type:=8)
    Set DestinationRange = Application.InputBox(Prompt:="Choose destination cell", Type:=8)
        
    CopyRange.Copy DestinationRange
    ```

    

- Returning an Array

  Sub ***ReturnArray***()

      Dim FilmLengths() As Variant
      FilmLengths = Application.InputBox(Prompt:="Choose lengths to convert", Type:=64)
      
      Dim LoopCounter As Long
      For LoopCounter = LBound(FilmLengths, 1) To UBound(FilmLengths, 1)
          
          FilmLengths(LoopCounter, 1) = Int(FilmLengths(LoopCounter, 1) / 60) & " hours " _
              & (FilmLengths(LoopCounter, 1) Mod 60) & " minutes"
      Next LoopCounter
      
      Dim ResultRange As Range
      Set ResultRange = _
          Application.InputBox(Prompt:="Choose where to put results", Type:=8)
          
      Set ResultRange = _
          Range(ResultRange, ResultRange.Offset(UBound(FilmLengths, 1) - 1, 0))
      
      ResultRange = FilmLengths

  End Sub