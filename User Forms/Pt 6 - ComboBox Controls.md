# Creating Drop Down Lists

[TOC]

![](C:\Users\13198\Documents\VBASource\Wise OWL\Image\UserFormPic\ComboBox Controls.jpg)

## Drawing and Formatting ComboBoxes

- FilmCertificate

  - Populating a Combo Box

    1. Using a Static List:    Lists!A2:A7

    2. Referring to Range Names: BBFCRatings

    3. Retrieving the Value of a Combo Box

       ```
       ActiveCell.Offset(0, 3).Value = FilmCertificate.Value
       ActiveCell.Offset(0, 3).NumberFormat = Range("E3").NumberFormat
       ```

    4. Retrieving the Choices in a Combo Box

       - The *MatchRequired* Property: 
       - *Style* Property:  2 - *fmStyleDropDownList*

    5. *Resetting the Formatting of a Combo Box

       Private Sub *FilmCertificate_AfterUpdate*()

           If FilmCertificate.Value <> "" Then
               FilmCertificate.BackColor = rgbWhite
               FilmCertificateLabel.ForeColor = Me.ForeColor
           End If

       End Sub

    6. Changing the Button and List Style:    *DropButtonStyle* & *ListStyle*

       

  - Using a Combo Box on a Form

  - *Validating a Combo Box

    Private Function *EverythingFilledIn*() As Boolean

        Dim ctl As MSForms.Control
        Dim AnythingMissing As Boolean
        
        EverythingFilledIn = True
        AnythingMissing = False
        
        For Each ctl In FilmDetailsFrame.Controls
            If TypeOf ctl Is MSForms.TextBox Or TypeOf ctl Is MSForms.ComboBox Then
                If ctl.Value = "" Then
                    ctl.BackColor = rgbPink
                    Controls(ctl.Name & "Label").ForeColor = rgbRed
                    If Not AnythingMissing Then ctl.SetFocus
                    AnythingMissing = True
                    EverythingFilledIn = False
                End If
            End If
        Next ctl

    End Function

    

- FilmCertificateLabel

  

## Alternatives to the Row Source

- *Setting the List Property

  Private Sub *UserForm_Initialize*()

      'FilmCertificate.RowSource = "BBFCRatings"

  ```
  '    FilmCertificate.List = _
  '        wsLists.Range("A2", wsLists.Range("A2").End(xlDown).End(xlToRight)).Value
  ```

      FilmCertificate.List = Range("BBFCRatings").Value

  End Sub

  

## Adding Items and Clearing Lists

- *The AddItem Method

  Private Sub *PopulateCertificates*()

      With FilmCertificate
          .AddItem "U"
          .AddItem "PG"
          .AddItem "12A"
          .AddItem "12"
          .AddItem "15"
          .AddItem "18"
      End With

  End Sub

- *Clearing and Resetting a List

  Private Sub *UseUKRatings_Click*()

      FilmCertificate.Clear
      
      '    FilmCertificate.List = _
      '        wsLists.Range("A2", wsLists.Range("A2").End(xlDown)).Value
      
      FilmCertificate.List = Range("BBFCRatings").Value

  End Sub

  Private Sub *UseUSRatings_Click*()

      FilmCertificate.Clear
      FilmCertificate.List = Range("MPAARatings").Value

  End Sub

## Setting Up a Multi-Column Combo Box 

- Setting the Multi-Column List

  `FilmCertificate.List =wsLists.Range("A2").End(xlDown).End(xlToRight)).Value`

  ...

- Working with Multi-Column Properties

  - *ColumnCount* [final 2]

  - *ColumnWidth*

  - *BoundColumn* :base 1 [final 1]

  - *TextColumn*: base 1 [final -1]

    ActiveCell.Offset(0, 4).Value = *FilmCertificate*.Text

    - Problems with TextColumn set 2:
      - Show Description not ...
      - Validation is not work
      - null empty

  - Referring to Columns by Index Number: base 0

    `ActiveCell.Offset(0, 3).Value = FilmCertificate.Column(0)`

  - **AddDataToList* subroutine

    Private Sub *AddDataToList*()

        wsFilms.Select
        Range("B2").End(xlDown).Offset(1, 0).Select
        
        ActiveCell.Value = FilmName.Value
        ActiveCell.Offset(0, 1).Value = FilmGross.Value
        ActiveCell.Offset(0, 1).NumberFormat = Range("C3").NumberFormat
        ActiveCell.Offset(0, 2).Value = FilmDate.Value
        ActiveCell.Offset(0, 2).NumberFormat = Range("D3").NumberFormat
        
        ActiveCell.Offset(0, 3).Value = FilmCertificate.Column(0, FilmCertificate.ListIndex)
        ActiveCell.Offset(0, 3).NumberFormat = Range("E3").NumberFormat
        
        ActiveCell.Offset(0, 4).Value = FilmCertificate.Column(1, FilmCertificate.ListIndex)
        
        MsgBox FilmName.Value & " was added to row " & ActiveCell.Row

    End Sub