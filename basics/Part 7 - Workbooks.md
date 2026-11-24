### Using Workbooks in VBA

[TOC]

#### Referring to Workbooks

- By Name

  Sub *ReferringToWorkbooksByName*()

  ```vbscript
  Workbooks("Book2.xlsx").Activate
  Workbooks("Top Movies 2012.xlsm").Activate
  ```

  End Sub

- By Number

  Sub *ReferringToWorkbooksByIndexNumber*()

  ```vbscript
  Workbooks(2).Activate
  Workbooks(1).Activate
  ```

  End Sub

- ActiveWorkbook

  Sub *UsingActiveWorkbook*()

  ```vbscript
  Workbooks("Book2.xlsx").Activate
  ActiveWorkbook.Close
  
  ActiveWorkbook.Close True
  ```

  End Sub

- ThisWorkbook

  Sub *UsingThisWorkbook*()

  ```vbscript
  Workbooks("Book2.xlsx").Activate
  
  ThisWorkbook.Close
  ```
  
  End Sub



#### Opening and Creating Workbooks

- Opening Existing Workbooks

  `Workbooks.Open ThisWorkbook.Path & "\Book2.xlsx"`

- Creating New Workbooks

   `Workbooks.Add "Top Movies 2012.xltm"`

#### Saving Workbooks

​	`Workbooks("Book2.xlsx").Save`

```vbscript
Workbooks.Add
ActiveWorkbook.Save
```

```vbscript
Workbooks.Add

ActiveWorkbook.SaveAs "C:\Test\Test workbook.xlsx"
```

- Changing the File Type

  ```vbscript
  Workbooks.Add
      ActiveWorkbook.SaveAs "C:\Test\Test workbook.xlsm", xlOpenXMLWorkbookMacroEnabled
  ```