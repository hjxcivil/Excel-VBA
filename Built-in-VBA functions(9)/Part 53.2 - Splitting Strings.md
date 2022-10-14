# VBA Functions for Splitting Strings

[TOC]

## Extracting Characters from the Left,Middle and Right

Sub *BasicStringSplittingFunctions*()

    Dim s As String
    
    s = Range("A15").Value 'I Am Legend
    
    Debug.Print s
    Debug.Print Right(s, 6)
    Debug.Print Left(s, 1) 
    Debug.Print Mid(s, 3, 2)
    
    Debug.Print Left$(s, 1) ' Left$ Return String While Left Variant 

    'Debug.Print Left(Null, 1) 'Null
    'Debug.Print Left$(Null, 1) 'Runtime-error

```
'Debug.Print LeftB(s, 1) ' empty?
'Debug.Print LeftB(s, 2) ' I
'Debug.Print LeftB(s, 6) ' I A
```


End Sub



## Finding the Position of a Character in a String

- *Instr* & *InstrRev*

- Calculating First and Last Names

  Sub *FindingCharacterPositions*() 'instr(1,"diY"," ")=0

      Dim s As String
      Dim FirstSpace As Long
      Dim LastSpace As Long
      
      s = Range("C43").Value 'J. J. Abrams
      
      Debug.Print s
      
      If s Like "* *" Then
          FirstSpace = InStr(1, s, " ") '3
          LastSpace = InStrRev(s, " ") '6
          Debug.Print Left(s, LastSpace - 1) 'J. J.
          'Debug.Print Right(s, Len(s) - FirstSpace)
          Debug.Print Mid(s, LastSpace + 1) 'Abrams
      Else
          Debug.Print s
      End If

  End Sub

## Splitting a Full Name into Two Parts

- Splitting Multiple Words

  Sub *LoopToSplitText*()
      
      Dim s As String
      Dim ThisSpace As Long
      Dim LastSpace As Long
      
      s = Range("A11").Value 'Harry Potter and the Order of the Phoenix
      
      ThisSpace = InStr(1, s, " ")
      
      If ThisSpace = 0 Then
          Debug.Print s
      Else
          Do While ThisSpace > 0
              Debug.Print Mid(s, LastSpace + 1, ThisSpace - LastSpace)
              
              LastSpace = ThisSpace
              
              ThisSpace = InStr(ThisSpace + 1, s, " ")
          Loop
          
          Debug.Print Mid(s, LastSpace + 1)
      End If

  End Sub

- Using the Split Function

  Sub *EasySplit*()

      Dim s As String
      Dim arr() As String
      Dim v As Variant
      
      s = Range("A11").Value
      
      arr = Split(s, " ")
      
      For Each v In arr
          Debug.Print v
      Next v

  End Sub

## Splitting Tab-Delimited Strings

- A Practical Example

  Sub *SplitTabDelimitedData*()

      Dim fso As New Scripting.FileSystemObject
      Dim ts As Scripting.TextStream
      Dim arr() As String
      Dim i As Integer, j As Integer
      
      Set ts = fso.OpenTextFile(Environ("UserProfile") & "\Desktop\HighGross.txt")
      
      Worksheets.Add
      
      Do Until ts.AtEndOfStream
          arr = Split(ts.ReadLine, vbTab)
          
          For i = LBound(arr) To UBound(arr)
              Cells(i + 1, j + 1).Value = arr(i)
          Next i
          
          j = j + 1
          
          Erase arr
      Loop
      
      ts.Close
      ActiveCell.CurrentRegion.EntireColumn.AutoFit

  End Sub