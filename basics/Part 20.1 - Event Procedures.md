# Create Event Procedures 

[TOC]

Examples of Events

- Opening and Closing Workbooks
- Printing and Saving Files
- Inserting Worksheets
- Selecting and Changing Cells

## Accessing the Events of an Object

- Workbook Events - Open

  Private Sub Workbook_Open()

  ​	`MsgBox "Hello " & Environ("Username"), , Date`

  End Sub

- Workbook Events - Close

  Private Sub Workbook_BeforeClose(Cancel As Boolean)

  ​	`MsgBox "Goodbye"`

  End Sub

  

## Adding Code to an Event



## Triggering an Event Procedure

Reopen to trigger Workbook_Open event

## A Note on Security



## Cancelling Events

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If Hour(Now) < 17 Then
        MsgBox "You're not leaving"
    End If

End Sub

- Other Cancellable Events

  Private Sub Workbook_BeforePrint(Cancel As Boolean)
      MsgBox "You can't print this file"
  
End Sub
  
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  
    MsgBox "You can't save this"
      Cancel = True
  
End Sub

## Disabling Events

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    Dim HowMany As Integer
    
    If TypeOf Sh Is Worksheet Then
        HowMany = InputBox("How many sheets would you like?")
        
        Application.EnableEvents = False
        Worksheets.Add Count:=HowMany - 1
        Application.EnableEvents = True
    End If

End Sub

## Worksheet Events - Selection Change

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    `Target.Interior.Color = vbYellow`

End Sub

- Restriction the Target Range

  Private Sub Worksheet_SelectionChange(ByVal Target As Range)

  ```
  If Target.Row <= 10 And Target.Column <= 5 Then
  	Target.Interior.Color = vbYellow
  End If
  ```

  End Sub

- Looping Over the Cells in the Target

      For Each SingleCell In Target
      
          If SingleCell.Row <= 10 And SingleCell.Column <= 5 Then
              SingleCell.Interior.Color = vbYellow
          End If
          
      Next SingleCell

- Counting the Cells in the Target

  ```
  If Target.Cells.CountLarge = 1 Then
      If Target.Row <= 10 And Target.Column <= 5 Then
      	Target.Interior.Color = vbYellow
      End If
  End If
  ```

  

## Worksheet Events - Change

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim SingleCell As Range
    
    If Target.Cells.CountLarge > 1000 Then Exit Sub
    
    For Each SingleCell In Target
    
    If SingleCell.Comment Is Nothing Then
        SingleCell.AddComment Now & " - " & SingleCell.Value & " - " & Environ("UserName")
    Else
        SingleCell.Comment.Text _
            vbNewLine & Now & " - " & SingleCell.Value & " - " & Environ("UserName") _
            , Len(SingleCell.Comment.Text) + 1 _
            , False
    End If
    
    SingleCell.Comment.Shape.TextFrame.AutoSize = True
    
    Next SingleCell

End Sub

## Events of Embedded Objects

Private Sub btnClearComments_Click()
    `Cells.ClearComments`

End Sub
