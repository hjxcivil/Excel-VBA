# How do I paste Excel data into an existing text box in PowerPoint?

[TOC]

## Referencing the PowerPoint Object Library

`Microsoft PowerPoint 16.0 Object Library`

## CopyDataIntoPowerpoint

Opening an Existing Presentation

​	Sub *CopyDataIntoPowerpoint*()



```
Dim ppt As PowerPoint.Application
Set ppt = New PowerPoint.Application
    
Dim pres As PowerPoint.Presentation
Set pres = ppt.Presentations.Open( _
    ThisWorkbook.Path & "\copy-excel-to-existing-powerpoint\MyPresentation.pptx")

Dim sl As PowerPoint.Slide:Set sl = pres.Slides(3)
Dim sh As PowerPoint.Shape:Set sh = sl.Shapes(2)

Worksheets(Curs).Range("A1").CurrentRegion.Copy
. . .
```

End Sub

 Pasting Excel Data into a Text Box

> sh.TextFrame.TextRange.Paste	

-  Pasting as Plain Text

  > ​	sh.TextFrame.TextRange.PasteSpecial ppPasteText 
  
  -  Pasting as RTF Text
  
    > sh.TextFrame.TextRange.PasteSpecial ppPasteRTF



- The TextFrame2 Object

> ​		sh.TextFrame2.TextRange.PasteSpecial msoClipboardFormatPlainText
>
> ​		sh.TextFrame2.TextRange.PasteSpecial msoClipboardFormatRTF

## Const Variables

​	

```
Private Const Curs As String = "copy-excel-to-existing-powerpoi"
```
