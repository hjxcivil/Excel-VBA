# The Select Case Statement

[TOC]

## Writing a Select Case Statement

Sub *TestingFilmLength*()

    Dim FilmName As String
    Dim FilmLength As Integer
    Dim FilmDescription As String
    
    Range("B11").Select
    
    FilmName = ActiveCell.Value
    FilmLength = ActiveCell.Offset(0, 2).Value
    
    'Select case statement goes here
    Select Case FilmLength
        Case Is < 100
            FilmDescription = "Short"
        Case Is < 120
            FilmDescription = "Medium"
        Case Is < 150
            FilmDescription = "Long"
        Case Else
            FilmDescription = "Epic"
    End Select
    
    MsgBox FilmName & " is " & FilmDescription

End Sub

- Testing for a Range of Numbers

  ```
  Select Case FilmLength
      Case 0 To 100
      	FilmDescription = "Short"
      Case 100 To 120
      	FilmDescription = "Medium"
      Case 120 To 150
      	FilmDescription = "Long"
      Case Else
      	FilmDescription = "Epic"
  End Select
  ```

- Testing for a List of Values

  ```
  Select Case FilmMonth
      Case 1, 2, 12
      	FilmSeason = "Winter"  
      Case 3, 4, 5
      	FilmSeason = "Spring"
      Case 6, 7, 8
      	FilmSeason = "Summer"
      Case 9, 10, 11
      	FilmSeason = "Autumn"
  End Select
  ```

## Nesting Select Case Statements

Sub *TestingFilmLength*()

    Dim FilmName As String
    Dim FilmLength As Integer
    Dim FilmDescription As String
    Dim FilmSeason As String
    Dim FilmMonth As Integer
    
    Range("B5").Select
    
    FilmName = ActiveCell.Value
    FilmLength = ActiveCell.Offset(0, 2).Value
    FilmMonth = Month(ActiveCell.Offset(0, 1).Value)
    
    Select Case FilmMonth
        Case 1, 2, 12
            Select Case FilmLength
                Case 0 To 100
                    FilmDescription = "Short Winter"
                Case 100 To 120
                    FilmDescription = "Medium Winter"
                Case 120 To 150
                    FilmDescription = "Long Winter"
                Case Else
                    FilmDescription = "Epic Winter"
            End Select
        Case 3, 4, 5
            Select Case FilmLength
                Case 0 To 100
                    FilmDescription = "Short Spring"
                Case 100 To 120
                    FilmDescription = "Medium Spring"
                Case 120 To 150
                    FilmDescription = "Long Spring"
                Case Else
                    FilmDescription = "Epic Spring"
            End Select
        Case 6, 7, 8
            FilmSeason = "Summer"
        Case 9, 10, 11
            FilmSeason = "Autumn"
    End Select
    
    MsgBox FilmName & " is " & FilmDescription

End Sub