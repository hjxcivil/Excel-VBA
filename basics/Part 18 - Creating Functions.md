# Returning Values from Procedures

[TOC]

## Writing a Simple Function

Function *CustomDate*() As String

    CustomDate = Format(Date, "dddd dd mmmm yyyy")

End Function

## Calling a Function

`Range("A1").Value = "Created on " & CustomDate`

## Creating Parameter

- Adding Parameters to a Function

  Function *CustomDate*(DateToFormat As Date) As String

  ​	   `CustomDate = Format(DateToFormat, "dddd dd mmmm yyyy")`

​	   End Function

- Call Function with Parameter

  `Range("A1").Value = "Created on " & CustomDate(#2/19/2014#)`

- Optional Parameters

  Function *CustomDate*(*DateToFormat* As Date, Optional *IncludeTime* As Boolean = False) As String

      If IncludeTime Then
          CustomDate = Format(DateToFormat, "dddd dd mmmm yyyy hh:mm:ss")
      Else
          CustomDate = Format(DateToFormat, "dddd dd mmmm yyyy")
      End If

  End Function

  `Range("A1").Value = "Created on " & CustomDate(Now, True)`

- Using Functions in Worksheets

  - In Worksheet : = *CustomDate*(...)
  - fx -> User Defined
  - Can be Refreshed by F9

## Rewriting Code to Use Functions

- Raw Code

  Sub *RateFilmsByLength*()

      Dim RunningTime As Integer
      
      Sheet1.Activate
      Range("A3").Select
      
      Do Until ActiveCell.Value = ""
          
         RunningTime = ActiveCell.Offset(0, 3).Value
         
         If RunTime < 100 Then
              ActiveCell.Offset(0, 4).Value = "Short"
          ElseIf RunTime < 150 Then
              ActiveCell.Offset(0, 4).Value = "Medium"
          ElseIf RunTime < 200 Then
              ActiveCell.Offset(0, 4).Value = "Long"
          Else
              ActiveCell.Offset(0, 4).Value = "Epic"
          End If
         
         ActiveCell.Offset(1, 0).Select
         
      Loop

  End Sub

- Rewrited Procedures to Use Functions

  Sub *RateFilmsByLength*()
      
      Sheet1.Activate
      Range("A3").Select
      
      Do Until ActiveCell.Value = ""
          
         ActiveCell.Offset(0, 4).Value = FilmLength(ActiveCell.Offset(0, 3).Value)
         
         ActiveCell.Offset(1, 0).Select
         
      Loop

  End Sub

- the FilmLength Function

  Function *FilmLength*(RunTime As Integer) As String

      If RunTime < 100 Then
          FilmLength = "Short"
      ElseIf RunTime < 150 Then
          FilmLength = "Medium"
      ElseIf RunTime < 200 Then
          FilmLength = "Long"
      Else
          FilmLength = "Epic"
      End If

  End Function