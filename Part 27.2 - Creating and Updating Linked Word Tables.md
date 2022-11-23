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

    







## Pasting Linked Excel Date

```
Worksheets(Curs).Range("A1:C5").Copy
wdApp.Selection.PasteExcelTable True, False, False
```



## Changing the Link Source - - - SINGLE - LINK

```
Dim wdField As Word.Field
Set wdField = wdDoc.Fields(1)
    
wdField.LinkFormat.SourceFullName = ThisWorkbook.FullName
```



## Creating and Updating Multiple Linked Tables



## Changing the Link Source using the Save Events



## Const Variables

```
Private Const SavePath = "C:\Users\13198\Desktop\vba-referencing-applications\linked-word-tables\Test\Linked Doc.docx"

Private Const SavePath2 = "C:\Users\13198\Desktop\vba-referencing-applications\linked-word-tables\Test\Multi Linked Doc.docx"

Private Const Curs = "linked-word-tables"
```

|item|qty|price||