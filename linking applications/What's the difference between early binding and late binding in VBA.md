# What's the difference between early binding and late binding in VBA

[TOC]

## Referencing an Object Library

`Microsoft Word 16.0 Object Library`

## Change The Early Binding

Sub WordReportEarlyBinding()

    Dim wd As Word.Application:Set wd = New Word.Application
    Dim doc As Word.Document:Set doc = wd.Documents.Add
    
    With wd.Selection
        .ParagraphFormat.Style = "Heading 1"
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .TypeText "Movie Report"
        .TypeParagraph
        .TypeParagraph
        .ParagraphFormat.Style = "Normal"
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    
       Worksheets(Curs).Range("A1").CurrentRegion.Copy
        .Paste
        .TypeParagraph
    
        Chart2.ChartArea.Copy
        .Paste
    
    End With
    
    doc.SaveAs2 ThisWorkbook.Path & "\Movie Report"
    wd.Quit

End Sub

## To Late Binding

- Defining Variables as Objects

  > Dim wd As Object:Dim doc As Object

- Using the CreateObject Function

  > Set wd = CreateObject("Word.Application")

- Replacing Constants with Numbers

  > .ParagraphFormat.Alignment = 1 'wdAlignParagraphCenter
  >
  > .ParagraphFormat.Alignment = 0 'wdAlignParagraphLeft

## Const Variables

​	

```
Private Const Curs As String = "early-binding-late-binding"
```
