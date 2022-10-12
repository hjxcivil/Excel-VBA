# Selecting Cells in VBA

[TOC]

## Selecting Cells by Absolute Position

- Selecting Single Cells 

  - Selecting Cells by Cell Reference

    `Range("A13").Select`

  - Referring to the Active Cell

    `ActiveCell.Value = 11`

  - Selecting Cells by Row and Column

    `Cells(13, 2).Select`

  - A Shorthand Way to Select Cells

     `[C13].Select`

  - Running The Code

    Sub *SelectSingleCellsByPosition*()

        Workbooks("Book2").Activate
        Worksheets("Sheet1").Activate
        
        Range("A13").Select
        ActiveCell.Value = 11
        
        Cells(13, 2).Select
        ActiveCell.Value = "The Lorax"
        
        [C13].Select
        ActiveCell.Value = #3/2/2012#   ' #2 Mar 2012# automatic change to this

    End Sub

  - Changing Cells without Selecting Them

    Sub *ChangeCellValuesWithoutSelecting*()

        Workbooks("Book3").Worksheets("Sheet2").Range("A14").Value = 12
        Workbooks("Book3").Worksheets("Sheet2").Range("B14").Value = "Wreck It Ralph"
        Workbooks("Book3").Worksheets("Sheet2").Range("C14").Value = #11/2/2012#

    End Sub

    

- Selecting Multiple Cells

  Sub *SelectMultipleCells*() 

      Range("A1:C1").Select 'ActiveCell only refer to one cell
      Selection.Interior.Color = rgbDarkBlue
      
      Range("A1:C1").Font.Color = rgbWhite
      
      [A1:C1].Font.Size = 14
      
      Range("A2", "C2").Interior.Color = rgbLightBlue
      
      Range(Cells(2, 1), Cells(2, 3)).Font.Color = rgbDarkBlue

  End Sub

- Creating Range Names in VBA:

  *Ctrl Shift F3*: Create Names from Selection

- Using Range Names

  Sub *ReferToRangeNames*()

      Range("ID").Font.Italic = True
      [Title].Font.Color = rgbDarkBlue

  End Sub

## Selecting Cells Relatively

- Finding the End of a List

- Moving Up, Down, Left, and Right

  Sub *AddFilmToEndOfList*()

      Worksheets("Sheet1").Activate
      
      Range("A1").End(xlDown).Offset(1, 0).Select
      
      ActiveCell.Value = ActiveCell.Offset(-1, 0).Value + 1
      ActiveCell.Offset(0, 1).Value = "Lincoln"
      ActiveCell.Offset(0, 2).Value = #11/9/2012# '#9 Nov 2012#    

  End Sub

- Selecting a List From the Top to Bottom

  Sub *SelectVariableCol*()

      Range("A3", Range("A1").End(xlDown)).Select
      Selection.Font.Italic = True
      
      Range("B3", Range("B2").End(xlDown)).Font.Color = rgbDarkBlue
      
      Range("A3", Range("A1").End(xlDown).End(xlToRight)).Select
      Selection.Interior.Color = rgbAliceBlue

  End Sub

- Selecting Entire Regions and Entire Columns

  - Copying and Pasting Cells

    Sub *CopyFilmList*() 

        Worksheets("Sheet1").Activate
        
        Range("A1").CurrentRegion.Copy
        
        Worksheets("Sheet3").Activate
        
        Range("A1").PasteSpecial
        Range("A1").PasteSpecial xlPasteColumnWidths

    End Sub

  - Copying Cells to a Destination

    Sub *CopyFilmListMethod2*()

        Worksheets("Sheet1").Activate
        
        Range("A1").CurrentRegion.Copy Worksheets("Sheet4").Range("A1")
        
        Worksheets("Sheet4").Activate
        
        'Range("B:B") & Range("B:C") & Columns("B")
        'Columns("B:C").Width = 20
        
        'Columns("B:C").AutoFit
        
        Range("B:C").EntireColumn.AutoFit

    End Sub



