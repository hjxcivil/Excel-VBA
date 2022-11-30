# How do I paste and resize a picture in Word with Excel VBA

[TOC]

## Referencing the PowerPoint Object Library

`Microsoft Word 16.0 Object Library`

Sub PasteWordPicAndResize()

    Dim wd As Word.Application
    Set wd = New Word.Application
    wd.Visible = True
    
    Dim doc As Word.Document
    Set doc = wd.Documents.Add
    ...

## PasteWordPicAndResize

- Typing Text into the Word Document 

  > With wd.Selection
  >         .Style = "Title"
  >         .TypeText "Top Grossing Films"
  >         .TypeParagraph
  >         .Style = "Normal"
  >
  > ​		Worksheets(Curs).Range("A1").CurrentRegion.Copy
  >
  > ...

  - Pasting Excel Data as a Word Table

    >
    > .Paste 'as a table

- The PasteSpecial Method

  > .PasteSpecial 'table default

  - Pasting Excel Data as a Picture 

    > .PasteSpecial DataType:=wdPasteEnhancedMetafile 

    - Setting the Placement of a Picture

      > .PasteSpecial ..., Placement:=wdInLine 'Default cann't drag to move
      >
      > .PasteSpecial ..., Placement:=wdFloatOverText 'can drag to move

  - Returning the Current Width and Height

    > Debug.Print doc.Shapes(1).Width, doc.Shapes(1).Height

  - Resizing a Floating Shape

    > doc.Shapes(1).Width = 400

  - Resizing an Inline Shape

    > doc.InlineShapes(1).Width = 400

## Const Variables

​	

```
Private Const Curs As String = "copy-excel-data-to-word"
```
