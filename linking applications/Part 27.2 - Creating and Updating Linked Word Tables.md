## Part 27.2 - Creating and Updating Linked Word Tables

![wddttb](../images/wddttb.PNG)

#### Creating a New Word Document

- Creating Sample Data

  ![LsSampleData](../images/LsSampleData.png)

  > [A1:C5]=RANDBETWEEN(1,100)

- Creating a Word Doc

  > Dim wdApp As New Word.Application

#### Pasting Linked Excel Data

- Using the Paste Excel Table Method

  > wdApp.Selection.PasteExcelTable True, False, False

- Checking The Document are Linked (F9 to refresh..)

  ![Linkbsc](../images/Linkbsc.PNG)

- Saving a Linked Doc

  > wdApp.ActiveDocument.SaveAs2 _
  >         ThisWorkbook.Path & "\Test\Linked Doc.docx"

#### Changing the Link Source

- Viewing Links in the Word Document

  ![EditLks](../images/EditLks.PNG)

- Changing the Linked Files Source & Write Code to Refresh

  ![image-20230530162406388](../../../AppData/Roaming/Typora/typora-user-images/image-20230530162406388.png)

#### Creating and Updating Multiple Linked Tables

- Create Multi Linked Table

  ![Multiwdtbl](../images/Multiwdtbl.PNG)

- Update Multi Word Links

  > Dim i As Integer
  >     For i = 1 To wdDoc.Fields.Count
  >         wdDoc.Fields(i).LinkFormat.SourceFullName = ThisWorkbook.FullName
  >     Next i

#### Changing the Link Source using the Save Events

- Workbook_BeforeSave (*Click Save*)

  ![bfsave](../images/bfsave.PNG)

- Workbook_AfterSave(*F12*)

  ![afsv](../images/afsv.PNG)