# Controlling Microsoft Word using Excel VBA

[TOC]

## Referencing the Word Object Library

`Dim wdApp As Word.Application`

## Creating a New Instance of Word

- Early binding

  Sub *CreateBasicWordReport*()

  

      Dim wdApp As Word.Application
      
      Set wdApp = New Word.Application
      
      wdApp.Visible = True
      wdApp.Activate

  End Sub

- Auto-Instance

  `Dim wdApp As New Word.Application`

- Late binding

  
  
      Dim wdApp As Object
      
      Set wdApp = CreateObject("Word.Application")

## Writing and Formatting Text

- Creating a New Document

  Sub *CreateBasicWordReport*()

  

      Dim wdApp As Word.Application
      
      Set wdApp = New Word.Application
      
      With wdApp
          .Visible = True
          .Activate
          
          .Documents.Add
      End With

  End Sub

- Typing and Formatting Text

  
  
      With wdApp
          .Visible = True:.Activate
          
          .Documents.Add
          
          With .Selection
              .ParagraphFormat.Alignment = wdAlignParagraphCenter
              .BoldRun
              .Font.Size = 14
              .TypeText "Top Movies of 2012"
              .BoldRun
              .TypeParagraph
              .Font.Size = 11
              .ParagraphFormat.Alignment = wdAlignParagraphLeft
              .TypeParagraph
          End With
          
      End With

## Copying Data into Word

    Range("A2", Range("A2").End(xlDown).End(xlToRight)).Copy
    
    wdApp.Selection.Paste

- Saving a Word Document

  `wdApp.ActiveDocument.SaveAs2 ...`

- Closing Documents and Quitting Word

  `wdApp.ActiveDocument.Close`

  `wdApp.Quit`

  `set wdApp=Nothing`

## Using Version-Specific Methods

Sub *CreateBasicWordReport*()

    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")
    
    With wdApp
        .Visible = True:.Activate
        
        .Documents.Add
        
        With .Selection
            .ParagraphFormat.Alignment = 1
            .BoldRun
            .Font.Size = 14
            .TypeText "Top Movies of 2012"
            .BoldRun
            .TypeParagraph
            .Font.Size = 11
            .ParagraphFormat.Alignment = 0
            .TypeParagraph
        End With
    
        Range("A2", Range("A2").End(xlDown).End(xlToRight)).Copy
        
        .Selection.Paste
        
        Dim FileExt As String
        FileExt = IIf(.Version <= 11, ".doc", ".docx")
        
        Dim SaveName As String
        SaveName = Environ("UserProfile") & "\Desktop\Movie Report " & _
            Format(Now, "yyyy-mm-dd hh-mm-ss") & FileExt
            
        If .Version <= 12 Then
            .ActiveDocument.SaveAs SaveName
        Else
            .ActiveDocument.SaveAs2 SaveName
        End If
        
        .ActiveDocument.Close
        .Quit
        
    End With
    
    Set wdApp = Nothing

End Sub

## Using Template

- Creating a Word Template
  1. Insert -> Links -> Bookmarts : *TableLocation*
  2. Ctrl + Enter : Generate new page
  3. Insert -> Links -> Bookmarts : *ChartLocation*
  4. Saveas Word Template(*.dotx) : Movie Report Template

- Creating Documents from Templates

  ```
  wdApp.Documents.Add "C:\Users\13198\Documents\自定义 Office 模板\Movie Report Template.dotx"
  ```

- Going to a Bookmark

  `wdApp.Selection.GoTo what:=-1, Name:="TableLocation"`

- Total

  Sub *CreateBasicWordReport*()

  
  
      Dim wdApp As Word.Application
      Set wdApp = CreateObject("Word.Application")
      
      With wdApp
          .Visible = True:.Activate
          
          .Documents.Add "C:\Users\13198\Documents\自定义 Office 模板\Movie Report Template.dotx"  
      
          Range("A2", Range("A2").End(xlDown).End(xlToRight)).Copy
      
          .Selection.GoTo what:=-1, Name:="TableLocation"
          .Selection.Paste
          
          Chart2.ChartArea.Copy
          
          .Selection.GoTo what:=-1, Name:="ChartLocation"
          .Selection.Paste
          
          Dim FileExt As String
          FileExt = IIf(.Version <= 11, ".doc", ".docx")
          
          Dim SaveName As String
          SaveName = Environ("UserProfile") & "\Desktop\Movie Report " & _
              Format(Now, "yyyy-mm-dd hh-mm-ss") & FileExt
              
           If .Version <= 12 Then
              .ActiveDocument.SaveAs SaveName
          Else
              .ActiveDocument.SaveAs2 SaveName
          End If
          
          .ActiveDocument.Close
          .Quit
          
      End With
      
      Set wdApp = Nothing
  
  End Sub