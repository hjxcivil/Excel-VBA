# Working with Sheets in VBA

[TOC]

## Referring to and Moving Between Sheets

- Activating a Worksheet

  `Worksheets("Sheet2").Activate`

  `Charts("Chart1").Activate`

  `Sheets("Sheet2").Activate`
  `Sheets("Chart1").Activate`

  

## Selecting Single and Multiple Sheets

- Selecting Rather than Activating

  `Worksheets("Sheet2").Select`

- Multiple Select

  ```
  Worksheets("Sheet1").Select
  Worksheets("Sheet2").Select False
  ```

  

## Sheet Names, Code Names and Index Numbers

- Problems with Using Sheet Names

- Using Worksheet Index Numbers

  `Worksheets(3).Select`
  `Sheets(4).Select`

- Using Sheet Code Names

  `Sheet2.Activate`
  `wsMovies.Activate`

  

## Manipulating Sheets

- Inserting Worksheets

  - `Worksheets.Add`
  - `Worksheets.Add before:=Sheets(1)`
     `Worksheets.Add after:=Sheets(Sheets.Count), Count:=3`

- Inserting Charts

  `Charts.Add after:=Charts("Chart1")` or
  `Sheets.Add Type:=XlSheetType.xlChart`

- Deleting Sheets

  Sub *DeleteSpecificSheets*()

      Application.DisplayAlerts = False
      Worksheets("Sheet5").Delete
      Sheets(Sheets.Count).Delete
      Application.DisplayAlerts = True

  End Sub

  `Charts.Delete`

- Copying Sheets

  - Basic Copying

    `wsMovies.Copy after:=Worksheets("Anything")`

  - Copying Sheets to a New Workbook

    `wsMovies.Copy`

  - Copying Sheets to an Open Workbook

    `wsMovies.Copy before:=Workbooks("Book4.xlsx").Sheets(1)`

- Moving Sheets

  `wsMovies.Move after:=Sheets(Sheets.Count)`

- Renaming Sheets

  `wsMovies.Name = "Something else"`

- Hiding and Showing Sheets

  `wsMovies.Visible = xlSheetVisible`
  `wsMovies.Visible = xlSheetVeryHidden 'xlSheetHidden`