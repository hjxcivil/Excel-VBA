# How do I copy Excel data into PowerPoint using VBA

[TOC]

## Referencing the PowerPoint Object Library

`Microsoft PowerPoint 16.0 Object Library`

## CopyDataIntoPowerpoint

- Creating a New Presentation 

      Dim ppt As PowerPoint.Application:Set ppt = New PowerPoint.Application
      Dim pres As PowerPoint.Presentation:Set pres = ppt.Presentations.Add

- Creating a  Blank Slide

  > Set sl = pres.Slides.Add(1, ppLayoutBlank)

- Creating Slides with a Custom Layout 

  > Dim cl As PowerPoint.CustomLayout
  > Set cl = pres.SlideMaster.CustomLayouts(7)
  >
  > Set sl = pres.Slides.AddSlide(1, cl)

- Copying and Pasting Excel Data into PowerPoint 
      

      Worksheets(Curs).Range("A1").CurrentRegion.Copy
      Dim sh As PowerPoint.ShapeRange
      Set sh = sl.Shapes.Paste 'as table

- Moving the Shape After Pasting

  > sh(1).Top = 20:sh(1).Left = 20

- Using the PasteSpecial Method

  - Pasting as Picture

    > Set sh = sl.Shapes.PasteSpecial(ppPasteEnhancedMetafile)

  - Pasting as Plain Text

    > Set sh = sl.Shapes.PasteSpecial(ppPasteText)

  - Pasting as RTF Text

    > Set sh = sl.Shapes.PasteSpecial(ppPasteRTF) 

  - Pasting as an OLE Object

    > Set sh = sl.Shapes.PasteSpecial(ppPasteOLEObject)

  - Linking an OLE Object to the Source File

    > Set sh = sl.Shapes.PasteSpecial(DataType:=ppPasteOLEObject, Link:=msoTrue)

## Const Variables

​	

```
Private Const Curs As String = "copy-excel-to-powerpoint"
```
