# Linking Excel Data in Word Documents

[TOC]

## Creating a Single - Linked Wd

- Creating a New Doc

  Sub ***CreateLinkedTable***()

  

  ```
  Dim wdApp As New Word.Application
  'wdApp.Visible = True
  wdApp.Documents.Add
  
  Worksheets(Curs).Range("A1:C5").Copy
  
  wdApp.Selection.PasteExcelTable True, False, False
  
  wdApp.ActiveDocument.SaveAs2 SavePath
  wdApp.Quit
  ```

  End Sub

  - Pasting Linked Excel Date

    > Worksheets(Curs).Range("A1:C5").Copy
    >
    > wdApp.Selection.PasteExcelTable True, False, False

- Changing the Link Source

  

  ```
  Dim wdField As Word.Field
  Set wdField = wdDoc.Fields(1)
  wdField.LinkFormat.SourceFullName = ThisWorkbook.FullName
  ```

  

## Creating and Updating Multiple Linked Tables

- Create CreateMultiLinkedTable

  Sub CreateMultiLinkedTable()

      Dim wdApp As New Word.Application
      
      'wdApp.Visible = True
      wdApp.Documents.Add
      
      Dim ws As Worksheet
      For Each ws In Worksheets
          If InStr(ws.Name, "linked-word-tables") Then
              Debug.Print ws.Name
              ws.Range("A1:C5").Copy
              wdApp.Selection.PasteExcelTable True, False, False
              wdApp.Selection.TypeParagraph
          End If
      Next ws
      
      wdApp.ActiveDocument.SaveAs2 SavePath2
      wdApp.Quit

  End Sub

- Changing the Link Source using the Save Events

  Sub UpdateMultiWordLinks()

      Dim wdApp As New Word.Application
      
      Dim wdDoc As Word.Document
      Set wdDoc = wdApp.Documents.Open(SavePath2)
      Dim i As Integer
      For i = 1 To wdDoc.Fields.Count
          wdDoc.Fields(i).LinkFormat.SourceFullName = ThisWorkbook.FullName
      Next i
      
      wdDoc.Save
      wdApp.Quit

  End Sub

## Const Variables

```
Private Const SavePath = "C:\Users\13198\Desktop\vba-referencing-applications\linked-word-tables\Test\Linked Doc.docx"

Private Const SavePath2 = "C:\Users\13198\Desktop\vba-referencing-applications\linked-word-tables\Test\Multi Linked Doc.docx"

Private Const Curs = "linked-word-tables"
```

|item|qty|price||