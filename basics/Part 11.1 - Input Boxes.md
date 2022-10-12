# Input Boxes

[TOC]

## Displaying & Customising an Input Box

```
InputBox "Please type in your name", "Personal Details", "Enter name here..."
```

## Capturing the Result

```
YourName = InputBox("Please type in your name", "Personal Details")
    
    If YourName = "" Then
        MsgBox "You didn't enter a name", vbExclamation
    Else
        MsgBox "Hello " & YourName
    End If
```

## Return Different Data Types

Sub ***CreateAFilm***() 'InputBox only return String type

```

    Dim FilmName As String
    FilmName = InputBox("Type in a film name") 'Gravity
    
    Dim FilmLength As Integer
    FilmLength = InputBox("Type in the length") '150'
    
    Dim strFilmDate As String
    strFilmDate = InputBox("Type in the release date") '8 Nov 2013
    
    If strFilmDate = "" Then
        MsgBox "You didn't enter a valid date":Exit Sub
    End If
    
    Dim datFilmDate As Date
    datFilmDate = CDate(strFilmDate)
    
    Range("B2").End(xlDown).Offset(1, 0).Select
    ActiveCell.Value = FilmName
    ActiveCell.Offset(0, 1).Value = datFilmDate
    ActiveCell.Offset(0, 2).Value = FilmLength
    
End Sub

```

~~Dim FilmDate As Date 'cann't store a empty string in a date variable~~
~~'FilmDate = InputBox("Type in the release date") 'Should be Date or can be translate to Date if cancel or close then cause runtime error~~
