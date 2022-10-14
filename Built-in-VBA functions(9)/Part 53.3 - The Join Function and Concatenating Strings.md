# Building String in VBA

[TOC]

## A Quick Recap of Concatenating

- Basic Concatenating

  Sub *BasicConcatenations*() 

      Dim s As String
      
      Range("A2").Select  '"Jurassic Park"
      
      s = ActiveCell.Value & "," & ActiveCell.Offset(0, 2).Value
      Debug.Print s 

  End Sub

- Adding and Converting Strings

  1. *if concatenation dif type with "+"  then cause run-time error*
  2. `s = CStr(ActiveCell.Value) + "," + CStr(ActiveCell.Offset(0, 3).Value)`

  

## Accumulating Values in String Variable

Sub *AccumulatingStrings*()

    Dim s As String
    Dim r As Range
    Dim Cols As Integer
    
    Cols = Range("A1").CurrentRegion.Columns.Count 
    
    For Each r In ActiveCell.Resize(1, Cols)
        s = s & r.Value & vbNewLine
    Next r
    
    s = Left(s, Len(s) - 1)
    Debug.Print s

End Sub

*Range("A1").CurrentRegion.EntireColumn Taller than Range("A1").CurrentRegion.Columns*

## The Join Function and Arrays

Sub *JoinFunction*()

    Dim s As String
    
    s = Join(Array("a", "b", "c"), vbTab)
    
    Debug.Print s

End Sub

## Transposing Arrays

- The Problem with Using a Range as an Array

  `s = Join(ActiveCell.Resize(1, Cols), vbTab)` 'Type mismatch

- Assigning a Range to an Array

      Dim arr()
      
      Cols = Range("A1").CurrentRegion.Columns.Count
      
      arr = ActiveCell.Resize(1, Cols).Value

- Transposing Rows and Columns

  Sub *JoiningRangeValues*()

      Dim s As String
      Dim Cols As Integer
      
      Cols = Range("A1").CurrentRegion.Columns.Count
      
      s = Join(Application.Transpose(Application.Transpose(ActiveCell.Resize(1, Cols).Value)), vbTab)
      
      Debug.Print s

  End Sub



## Writing Data to Text Files

- Generating a Text File

  Sub *ListActionFilms*()

      Dim fso As New Scripting.FileSystemObject
      Dim ts As Scripting.TextStream
      Dim r As Range
      Dim s As String
      Dim Cols As Integer
      
      Cols = Range("A1").CurrentRegion.Columns.Count
      
      Set ts = _
          fso.OpenTextFile(Environ("UserProfile") & "\Desktop\Action.txt", ForAppending, True)
      
      ...
      
      ts.Close

  End Sub

- Writing String to a Text File

      ...
      For Each r In Range("A2", Range("A1").End(xlDown))
      	If LCase(r.Offset(0, 5).Value) = "action" Then
              s = Join(Application.Transpose(Application.Transpose(r.Resize(1, Cols).Value)), vbTab)
              ts.WriteLine s
          End If
      
      Next r
      ...

- Generating and Writing to Multiple Text Files

  Sub *CreateGenreFiles*()

      Dim fso As New Scripting.FileSystemObject
      Dim ts As Scripting.TextStream
      Dim r As Range
      Dim s As String
      Dim Cols As Integer
      Dim FolPath As String
      Dim Genre As String
      
      Cols = Range("A1").CurrentRegion.Columns.Count
      
      FolPath = Environ("UserProfile") & "\Desktop\Genres"
      
      If Not fso.FolderExists(FolPath) Then fso.CreateFolder FolPath
      
      For Each r In Range("A2", Range("A1").End(xlDown))
      
          Genre = r.Offset(0, 5).Value
          
          Set ts = _
              fso.OpenTextFile(FolPath & "\" & Genre & ".txt", ForAppending, True)
          
          s = Join(Application.Transpose(Application.Transpose(r.Resize(1, Cols).Value)), vbTab)
          ts.WriteLine s
          
          ts.Close
      Next r

  End Sub