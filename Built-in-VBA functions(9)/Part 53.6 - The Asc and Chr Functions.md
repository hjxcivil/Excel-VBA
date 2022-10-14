# Wroking with Character Codes in VBA

[TOC]

## The Asc and Chr Functoins

Sub *TheBasics*()

    Debug.Print Asc("A") '65
    Debug.Print Asc("ABC") '65
    
    Debug.Print Chr(65) 'A
    Debug.Print Chr(32) '[space]

End Sub

## Creating an ASCII Reference Table

- List Characters

  Sub *ListCharacters*() 

      Dim i As Integer
      
      Worksheets.Add
      
      For i = 0 To 255
          Cells(i + 1, 1).Value = i
          Cells(i + 1, 2).Value = Chr(i)
      Next i

  End Sub

- Listing ASCII Codes from a String

  Sub *ListCodes*() 

      Dim i As Integer
      Dim s As String
      
      s = ActiveCell.Value
      
      For i = 1 To Len(s)
          Debug.Print Asc(Mid(s, i, 1))
      Next i

  End Sub

## ASCII Control Characters

- *Chr [0-32] is ControlCharacters*

- *VbCrLf ,VbNewLine*

- *Chr(13) => VbCr Chr(10) => VbLf Chr(9) => VbTab*

- *Chr(13):CR ,Chr(9) & vbTab not performance in Sheet cell*

  Sub *ControlCharacters*()

      Const s1 As String = "Wise"
      Const s2 As String = "Owl"
      
      Debug.Print s1 & Chr(13) & Chr(10) & s2
      Debug.Print s1 & vbCrLf & s2
      Debug.Print s1 & vbNewLine & s2
      Debug.Print s1 & vbCr & s2
      Debug.Print s1 & vbLf & s2
      
      Debug.Print s1 & Chr(9) & s2
      Debug.Print s1 & vbTab & s2
      
      Worksheets.Add
      
      Range("A1").Value = s1 & Chr(13) & Chr(10) & s2
      Range("A2").Value = s1 & vbCrLf & s2
      Range("A3").Value = s1 & vbNewLine & s2
      Range("A4").Value = s1 & vbCr & s2
      Range("A5").Value = s1 & vbLf & s2
      
      Range("A6").Value = s1 & Chr(9) & s2
      Range("A7").Value = s1 & vbTab & s2

  End Sub

## Unicode Characters

- Unicode < 0 or > 255

  Sub *Unicode*() 
          

      ActiveCell.Value = ChrW(&H2776)
      Debug.Print AscW(ActiveCell.Value) '10102
      ActiveCell.Offset(1, 0).Value = ChrW(10102)

  End Sub

- Inserting a Unicode Character

  *Insert -> Symblol   : ❶ [Normal text, Dingbats,2776] -> [a1]*

- Returning a Unicode Character Code

  `?asc([a1]) => 63 , ?ascW([a1]) => 10102`

- Printing Unicode Characters

  `?chrw(10102) -> ? ;  ActiveCell.Value = chrw(10102) =>❶` 

- Using Hex Numbers

  ```
  ??hex(10102) -> 2776
  
  ActiveCell.Value = chrw(10102) ❶ 
  
  ActiveCell.Value = chrw(&H2776) ❶
  ```

   

- Listng Unicode Characters

  Sub *ListUnicodeCharacters*()

      Dim n As Long
      
      Application.ScreenUpdating = False
      
      Worksheets.Add
      
      For n = -32768 To 65535
          Cells(n + 32769, 1).Value = ChrW(n)
          Cells(n + 32769, 2).Value = n
          Cells(n + 32769, 3).Value = Hex(n)
      Next n
      
      Application.ScreenUpdating = True

  End Sub

## An Obsolete Practical Example

- Creating an Enumeration

  Enum *CharType*
      

  ```
  Unicode = 0  'Const Unicode as Long = 0 if no enum
  Uppercase = 1
  LowerCase = 2
  Number = 3
  Other = 4
  ```

  End Enum

- Looping Over a Range of Cells

  Sub *ListingCharacterTypes*()

      Dim r As Range
      Dim n As Long, nRows As Long
      
      Sheet1.Select
      
      Set r = Range("A2", Range("A1").End(xlDown))
      nRows = r.Rows.Count
      
      For n = 1 To nRows
          ...
      Next n

  End Sub

- Looping Over a Range of Characters

- A Function to Calculate the Character Type

  Function *CharacterType*(Character As String) As *CharType*

      Select Case AscW(Character)
          Case Is < 0
              CharacterType = Unicode
          Case Is > 255
              CharacterType = Unicode
          Case 48 To 57
              CharacterType = Number
          Case 65 To 90
              CharacterType = Uppercase
          Case 97 To 122
              CharacterType = LowerCase
          Case Else
              Character = Other
      End Select

  End Function

- Copmlete

  Sub *ListingCharacterTypes*()

      Dim r As Range
      Dim n As Long, nRows As Long
      Dim s As String
      Dim i As Integer
      Dim ct As CharType
      Dim arr()
      
      Sheet1.Select
      
      Set r = Range("A2", Range("A1").End(xlDown))
      nRows = r.Rows.Count
      
      ReDim arr(1 To nRows, 0 To 4)
      
      For n = 1 To nRows
      
          s = r.Cells(n, 1).Value
          
          For i = 1 To Len(s)
              ct = CharacterType(Mid(s, i, 1))
              'Debug.Print ct
              arr(n, ct) = arr(n, ct) + 1
          Next i
          
      Next n
      
      r.Offset(0, 1).Resize(nRows, 5) = arr

  End Sub