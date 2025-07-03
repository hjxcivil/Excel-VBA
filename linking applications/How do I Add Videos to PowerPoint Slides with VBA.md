### How do I Add Videos to PowerPoint Slides with VBA?

- #### Add online or local picture to slide

  ![PixPin_2025-07-04_05-14-34](../images/PixPin_2025-07-04_05-14-34.png)

- #### Add Videos to slide

  ![PixPin_2025-07-04_05-47-44](../images/PixPin_2025-07-04_05-47-44.png)



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

  ![Uplks](../images/Uplks.PNG)

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