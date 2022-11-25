# Controlling Microsoft Word using Excel VBA

[TOC]

## Referencing the PowerPoint Object Library

`Dim ppApp As PowerPoint.Application`



## Creating a New Instance of PowerPoint

- Early binding

  Sub *CreateNewPresentation*()

  

      Dim ppApp As PowerPoint.Application
      Set ppApp = New PowerPoint.Application
      
      ppApp.Visible = True
      ppApp.Activate
  
  End Sub
  
- Auto-Instance

  `Dim ppApp As New PowerPoint.Application`

- Late binding

  
  
      Dim wdApp As Object
      
      Set wdApp = CreateObject("PowerPoint.Application")

## Writing and Formatting Text

- Creating a New Presentation

  Sub *CreateNewPresentation*()

  

      Dim ppApp As PowerPoint.Application
      Set ppApp = New PowerPoint.Application
      
      ppApp.Visible = True
      ppApp.Activate
      
      Dim ppPres As PowerPoint.Presentation
      Set ppPres = ppApp.Presentations.Add
      . . .
  
  End Sub
  
- Creating a Title Slide

  > ​	Dim ppSlide As PowerPoint.Slide
  > ​    Set ppSlide = ppPres.Slides.Add(1, ppLayoutTitle)
  
  - Adding Text to Textboxes
  
    > ​	ppSlide.Shapes(1).TextFrame.TextRange.Text = "Movie Presentation"
    > ​    ppSlide.Shapes(2).TextFrame.TextRange.Text = "By Wise Owl"
  
- Copying an Excel Range into Power Point

  ```
  Set ppSlide = ppPres.Slides.Add(2, ppLayoutBlank)
  ppSlide.Select
  
  Sheet1.Range("A1").CurrentRegion.Copy
  
  'generate an embeded execl obj
  'ppSlide.Shapes.PasteSpecial ppPasteOLEObject 
  
  'generate a powerpoint table
  ppSlide.Shapes.Paste.Select [use select to fixed lower bundle version] 
  
  
  ```

  - Modifying the position

    > ppSlide.Shapes(1).Width = ppPres.PageSetup.SlideWidth
    > ppSlide.Shapes(1).Left = 0
    >     
    >
    > ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    >
    > OR
    >
    > ppSlide.Shapes(1).Top = _
    >         (ppPres.PageSetup.SlideHeight / 2) - (ppSlide.Shapes(1).Height / 2)

  - Adding & formatting a Custom Textbox

    
        Dim ppTextbox As PowerPoint.Shape
        	Set ppTextbox = ppSlide.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, 0, 20, ppPres.PageSetup.SlideWidth, 60)
                
        With ppTextbox.TextFrame
            .TextRange.Text = "List of Current Films"
            .TextRange.ParagraphFormat.Alignment = ppAlignCenter
            .TextRange.Font.Size = 26
            .TextRange.Font.Name = "Calibri"
            .VerticalAnchor = msoAnchorMiddle
        End With

- Copying an Excel Chart into Power Point

  

      Set ppSlide = ppPres.Slides.Add(3, ppLayoutBlank)
      ppSlide.Select
      
      Chart1.ChartArea.Copy
      ppSlide.Shapes.Paste.Select

  - Align Middle and Center

    > If ppSlide.Shapes(1).Width > ppPres.PageSetup.SlideWidth Then
    >         ppSlide.Shapes(1).Width = ppPres.PageSetup.SlideWidth
    >     End If
    >     
    >
    > ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
    > ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoTrue



- Applying a Template to a Presentation

  > ppPres.ApplyTemplate TemplatePath
  > ppPres.ApplyTheme ThemePath

- Saving and Closing a Presentation
      

      Dim FileExt As String
      FileExt = IIf(ppApp.Version <= 11, ".ppt", ".pptx")
      ppPres.SaveAs Environ("UserProfile") & "\Desktop\Movie Pres " & _
              Format(Now, "yyyymmddhhmmss") & FileExt
      ppPres.Close:ppApp.Quit
      Set ppApp = Nothing





## Const Variables

​	

```
Private Const TemplatePath As String = "C:\Program Files (x86)\Microsoft 		Office\root\Templates\2052\Pitchbook.potx"
Private Const ThemePath As String = "C:\Users\13198\Desktop\vba-referencing-		applications\powerpoint-presentations\Ion.thmx"
```





## Converting to Use Late Binding

Sub ***CreateNewPresentationWithCreateObject***()

```
Dim ppApp As Object
Set ppApp = CreateObject("PowerPoint.Application")

Dim ppPres As Object
Set ppPres = ppApp.Presentations.Add

On Error Resume Next
ppPres.ApplyTemplate TemplatePath
On Error GoTo 0

Dim ppSlide As Object
Set ppSlide = ppPres.Slides.Add(1, 1)

ppSlide.Shapes(1).TextFrame.TextRange.Text = "Movie Presentation"
ppSlide.Shapes(2).TextFrame.TextRange.Text = "By Wise Owl"

Set ppSlide = ppPres.Slides.Add(2, 12)
ppSlide.Select

Sheet1.Range("A1").CurrentRegion.Copy
ppSlide.Shapes.PasteSpecial 10
ppSlide.Shapes(1).Width = ppPres.PageSetup.SlideWidth
ppSlide.Shapes(1).Left = 0

ppSlide.Shapes(1).Top = _
    (ppPres.PageSetup.SlideHeight / 2) - (ppSlide.Shapes(1).Height / 2)

Dim ppTextbox As Object
Set ppTextbox = ppSlide.Shapes.AddTextbox( _
    msoTextOrientationHorizontal, 0, 20, ppPres.PageSetup.SlideWidth, 60)

With ppTextbox.TextFrame
    .TextRange.Text = "List of Current Films"
    .TextRange.ParagraphFormat.Alignment = 2
    .TextRange.Font.Size = 26
    .TextRange.Font.Name = "Calibri"
    .VerticalAnchor = 3
End With

Set ppSlide = ppPres.Slides.Add(3, 12)
ppSlide.Select

Chart1.ChartArea.Copy
ppSlide.Shapes.Paste.Select

If ppSlide.Shapes(1).Width > ppPres.PageSetup.SlideWidth Then
    ppSlide.Shapes(1).Width = ppPres.PageSetup.SlideWidth
End If

ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
ppApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoTrue

Dim FileExt As String
FileExt = IIf(ppApp.Version <= 11, ".ppt", ".pptx")
    
ppPres.SaveAs Environ("UserProfile") & "\Desktop\Movie Pres " & _
        Format(Now, "yyyymmddhhmmss") & FileExt

ppPres.Close
ppApp.Quit

Set ppApp = Nothing
```

End Sub