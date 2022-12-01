# How do I copy Word tables into Excel using VBA

[TOC]

## Referencing the PowerPoint Object Library

`Microsoft Word 16.0 Object Library`

Sub ImportWrodTables()

    Dim wd As Word.Application
    Set wd = New Word.Application
    wd.Visible = True
    
    ...

- Opening a Word Document

  > Dim doc As Word.Document
  > Set doc = wd.Documents.Open(FilePath)

- Looping Through the Tables Collection

      ...
      Dim tbl As Word.Table,ws As Worksheet
      For Each tbl In doc.Tables
          tbl.Range.Copy
          Set ws = ThisWorkbook.Worksheets.Add
          ws.PasteSpecial
          ...
      Next tbl
      doc.Close:wd.Quit

- Paste Special Format Options

  - Pasting as Plain Text

    > ws.PasteSpecial "Text"

  -  Formatted Text

    > ws.PasteSpecial "HTML"

- Changing the Column Widths

  > ws.Range("A1").CurrentRegion.EntireColumn.AutoFit

- ***ImportWholeDoc***

  > doc.Range.Copy ws.Paste

- Others: 

> ​		tbl.Range.Information(wdActiveEndPageNumber)

## Const Variables

​	

```
Private Const FilePath As String = "C:\Users\13198\Desktop\vba-referencing-applications\copy-word-table-to-excel\Movies.docx"

```
