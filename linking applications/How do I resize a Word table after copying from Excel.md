# How do I resize a Word table after copying from Excel

[TOC]

## Referencing the Word Object Library 

`Microsoft Word 16.0 Object Library`

## Resize a Word table Single Sheet

Sub CopyTableToWord()

    Dim wd As Word.Application
    Set wd = New Word.Application
    wd.Visible = True
    
    Dim doc As Word.Document
    Set doc = wd.Documents.Add
    ...

- Copying Excel Data to Word : Region is too large to display in word

  > ​	Sheet1.Range("A1").CurrentRegion.Copy
  > ​    wd.Selection.Paste

  - Using the Autofit Method 

    > Dim tbl As Word.Table
    >     Set tbl = doc.Tables(1)
    >     tbl.Columns.AutoFit

  - Making Columns the Same Width 

    > tbl.Columns.DistributeWidth

  - Setting the Preferred Width of the Table 

    > ​    tbl.PreferredWidthType = wdPreferredWidthPercent
    > ​    tbl.PreferredWidth = 50
    > ​    tbl.Columns.DistributeWidth

  - Reducing the Table Width 

    > tbl.AllowAutoFit = False
    >     tbl.Columns.DistributeWidth
    >     tbl.PreferredWidthType = wdPreferredWidthPercent
    >     tbl.PreferredWidth = 50

## Resize a Word table Multi Sheets

- Looping Through the Worksheets Collection

        Dim ws As Worksheet
        For Each ws In Worksheets  
          If ws.Name Like "Y20??" Then
              ws.Range("A1").CurrentRegion.Copy
              wd.Selection.Paste
      
              Dim tbl As Word.Table
              Set tbl = doc.Tables(doc.Tables.Count)
      
              tbl.AllowAutoFit = False
              tbl.Columns.DistributeWidth
              tbl.PreferredWidthType = wdPreferredWidthPercent
              tbl.PreferredWidth = 100
              ...
          End If
      Next ws

- Adding Page Breaks

  > ...
  >
  > tbl.PreferredWidth = 100
  >       wd.Selection.InsertBreak wdPageBreak
  > End If

  - Last Page Need No PageBreak

        ...
        Next ws
        wd.Selection.TypeBackspace:wd.Selection.TypeBackspace