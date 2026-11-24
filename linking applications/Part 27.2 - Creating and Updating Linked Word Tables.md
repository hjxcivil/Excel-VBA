### Part 27.2 - Creating and Updating Linked Word Tables

![wddttb](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014537483.PNG)

#### Creating a New Word Document

- Creating Sample Data

  ![LsSampleData](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014550505.png)

  > [A1:C5]=RANDBETWEEN(1,100)

- Creating a Word Doc

  > Dim wdApp As New Word.Application

#### Pasting Linked Excel Data

- Using the Paste Excel Table Method

  > wdApp.Selection.PasteExcelTable True, False, False

- Checking The Document are Linked (F9 to refresh..)

  ![Linkbsc](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014606833.PNG)

- Saving a Linked Doc

  > wdApp.ActiveDocument.SaveAs2 _
  >         ThisWorkbook.Path & "\Test\Linked Doc.docx"

#### Changing the Link Source

- Viewing Links in the Word Document

  ![EditLks](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014619483.PNG)

- Changing the Linked Files Source & Write Code to Refresh

  ![Uplks](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014632351.PNG)

#### Creating and Updating Multiple Linked Tables

- Create Multi Linked Table

  ![Multiwdtbl](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014651182.PNG)

- Update Multi Word Links

  > Dim i As Integer
  >     For i = 1 To wdDoc.Fields.Count
  >         wdDoc.Fields(i).LinkFormat.SourceFullName = ThisWorkbook.FullName
  >     Next i

#### Changing the Link Source using the Save Events

- Workbook_BeforeSave (*Click Save*)

  ![bfsave](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014711851.PNG)

- Workbook_AfterSave(*F12*)

  ![afsv](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014723079.PNG)