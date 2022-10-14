# Changing the Case of Strings in VBA

[TOC]

## Why Case is Important

Sub *ComparingStringCase*()

    Dim s As String
    Dim r As Range
    
    For Each r In Range("A2",Range("A1").End(xlDown))
        
        s = r.Offset(0, 5).Value
        
        If s = "Action" Then Debug.Print r.Value
    
    Next r

End Sub

## The Option Compare Statement

`Option Compare Text 'Binary default`

## Comparing Strings with UCase, LCase and StrComp

- `s = LCase(r.Offset(0, 5).Value)`

- `s = UCase(r.Offset(0, 5).Value)`

- `If StrComp(s, "acTioN", vbTextCompare) = 0 Then`

- Creating a Useful Example

  Sub *ComparingStringCase*()

      Dim s As String
      Dim r As Range
      Dim ws As Worksheet
      
      Application.ScreenUpdating = False
      
      Set ws = Worksheets.Add
      Sheet1.Range("A1").EntireRow.Copy ws.Range("A1")
      Range("A2").Select
      
      For Each r In Sheet1.Range("A2", Sheet1.Range("A1").End(xlDown))
          
          s = r.Offset(0, 5).Value
      
          If StrComp(s, "acTioN", vbTextCompare) = 0 Then
              r.EntireRow.Copy ActiveCell
              ActiveCell.Offset(1, 0).Select
          End If
      
      Next r
      
      Range("A1").Select
      ActiveCell.CurrentRegion.EntireColumn.AutoFit
      
      Application.ScreenUpdating = True

  End Sub

## Converting Strings to Upper, Lower and Proper Case

*LCase*(s):   `StrConv(s, vbLowerCase)`
*UCase*(s):  `StrConv(s, vbUpperCase)`
`WorksheetFunction.Proper(s)   StrConv(s, vbProperCase)`

Sentence Case: `UCase(Left(s, 1)) & LCase(Mid(s, 2))`

Sub *CreateSentenceCase*()

    Dim s As String
    Dim sentences() As String
    Dim i As Integer
    
    s = "Harry Potter and the Order of the Phoenix. Harry Potter and the Goblet of Fire. Harry Potter and the Prisoner of Azkaban."
    
    sentences = Split(s, ".")
    
    For i = LBound(sentences) To UBound(sentences)
        If sentences(i) <> "" Then
            s = Trim(sentences(i))
            s = UCase(Left(s, 1)) & LCase(Mid(s, 2))
            sentences(i) = s
        End If
    Next i
    
    s = Trim(Join(sentences, ". "))
    
    Debug.Print s

End Sub

## Creating a Sentence Case Function

Function *SentenceCase*(s As String) As String

    Dim sentences() As String
    Dim i As Integer
        
    sentences = Split(s, ".")
    
    For i = LBound(sentences) To UBound(sentences)
        If sentences(i) <> "" Then
            s = Trim(sentences(i))
            s = UCase(Left(s, 1)) & LCase(Mid(s, 2))
            sentences(i) = s
        End If
    Next i
    
    s = Trim(Join(sentences, ". "))
    
    SentenceCase = s

End Function

## Creating a Toggle Case Function

Sub *ToggleCase*()

    Dim s1 As String, s2 As String
    Dim c As String * 1
    Dim n As Long
    
    s1 = "aBCdeF"
    
    For n = 1 To Len(s1)
        c = Mid(s1, n, 1)
        
        If StrComp(c, UCase(c), vbBinaryCompare) = 0 Then
            c = LCase(c)
        Else
            c = UCase(c)
        End If
        
        s2 = s2 & c
        
    Next n
    
    Debug.Print s1
    Debug.Print s2

End Sub

Function *ToggleCase*(s1 As String) As String

    Dim s2 As String
    Dim c As String * 1
    Dim n As Long
    
    For n = 1 To Len(s1)
        c = Mid(s1, n, 1)
        
        If StrComp(c, UCase(c), vbBinaryCompare) = 0 Then
            c = LCase(c)
        Else
            c = UCase(c)
        End If
        
        s2 = s2 & c
        
    Next n
    
    ToggleCase = s2

End Function