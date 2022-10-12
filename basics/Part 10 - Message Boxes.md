# Displaying Messages on Screen

[TOC]

## The MsgBox Function

`MsgBox "I like Pizza!"`

## Customising Message Boxes

`MsgBox "I like Pizza!", vbInformation, "Food message"`

- Using Named Parameters

  `MsgBox prompt:="I like Pizza!", Buttons:=vbInformation, Title:="Food message"`

## Concatenating Strings

`MsgBox "The date is " & Date & ". The weather is rainy."`

## Adding Extra Lines to Messages

`MsgBox "The date is " & Date & "." & vbNewLine & "The weather is rainy."`

## Displaying Values from Cells

Sub *MovieMessage*()

    Range("B12").Select
    
    MsgBox ActiveCell.Value & " was released on " & ActiveCell.Offset(0, 1).Value

End Sub

## Asking Questions with Message Boxes

Sub *SimpleMessage*()

    Dim ButtonClicked As VbMsgBoxResult
    
    MsgBox prompt:="I like Pizza!", Buttons:=vbInformation, Title:="Food message"
    
    ButtonClicked = MsgBox("Do you like pizza?", vbQuestion + vbYesNo, "Food Question")
    
    If ButtonClicked = vbYes Then
        MsgBox "Yes, pizza's are great!", vbExclamation
    Else
        MsgBox "Why not? Pizza's are great!", vbCritical
    End If

End Sub