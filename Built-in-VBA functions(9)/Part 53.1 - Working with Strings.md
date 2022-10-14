# Working with Text in VBA

[TOC]

## Declaring String Variables

Sub *BasicStrings*()

`Dim s$`

    Dim s As String
    
    s = "anything you like up to about 2.147 billion characters"
    s = ThisWorkbook.Path
    s = Application.Version
    s = ActiveSheet.Name
    s = ActiveCell.Address
End Sub

## Fixed Length and Variable Length

`Dim s As String * 5`

## Converting Values to Strings

- Converting Values to Strings Implicitly

  ```
  Dim s As String
  s = 123          => "123"
  s = #2/19/2017#  => "2017/2/19"
  ```

- Converting Values to Strings Explicitly

  ```
  s = CStr(ActiveCell.Row)
  s = CStr(ThisWorkbook.Sheets.Count)
  s = CStr(Date)
  ```

## Dealing with Nulls

​	<!--Both s = CStr(Null) and s=Null will cause error-->

​	`s = IIf(IsNull(Null), "", "reference to field!")`

## Concatenation

- Basic

  ```
  s1 = "a"
  s2 = "b"
  s3 = s1 & s2 's1 + s2 also result in "ab"
  ```

- Concatenate Cells

  ```
  s1 = ActiveCell.Value '"Pearl Harbor"
  s2 = ActiveCell.Offset(0, 3).Value '"183"
  s3 = s1 & ", " & s2 
  ```

- & vs. +

  Sub *ConcatenationOperators*() 

      "a" & "b"   => "ab"
      "a" + "b" 	=> "ab"
      "2" & "2"   => "22"
      "2" + "2"   => "22"
      2 + 2  	=> 4
      "2" + 2  	=> 4
      2 & 2   	=>"22"
      
      ActiveCell.Value & ", " & ActiveCell.Offset(0, 3).Value '"Pearl Harbor, 183"
      ActiveCell.Value + ", " & ActiveCell.Offset(0, 3).Value '"Pearl Harbor, 183"
      ActiveCell.Value + ", " + ActiveCell.Offset(0, 3).Value ' Type Mismatch error
      
      ActiveCell.Value + ", " + CStr(ActiveCell.Offset(0, 3).Value) '"Pearl Harbor, 183"
      
      Range("A128").Select
      Debug.Print CStr(ActiveCell.Value) + ", " + CStr(ActiveCell.Offset(0, 3).Value) '"300, 117"
      Debug.Print ActiveCell.Value & ", " & ActiveCell.Offset(0, 3).Value '"300, 117"

  End Sub

## String Constants(vbTab, vbNewLine, vbCrLf)

- Tab-Delimited

  ```
  For Each r In Range(ActiveCell, ActiveCell.End(xlToRight))
  	s = s & r.Value & vbTab
  Next r
  ```

- NewLine

  ```
  For Each r In Range(ActiveCell, ActiveCell.End(xlToRight))
  	s = s & r.Value & IIf(r.Offset(0, 1).Value = "", "", vbNewLine)
  Next r
  ```

- VbCr may a little Different from others[vbNewLine,vbCrLf,vbLf]

  `s3 = s1 & Chr(13) & Chr(10) & s2`

## Comparing Strings

- Case Sensitive by default

- `Option Compare Text`

- Lcase

- Comparing Strings in a Loop

  Sub *WildcardComparisons*() 

      Dim r As Range
      
      For Each r In Range("A2", Range("A1").End(xlDown))
          
          If LCase(r.Value) = "king kong" Then
              Debug.Print r.Value & ", " & r.Offset(0, 1).Value
          End If
      Next r

  End Sub

## Wildcards and Pattern Matching

-  *?[]!#-
          

      `If LCase(Left(r.Value, 1)) = "k" Then`
      `If LCase(r.Value) Like "k*" Then`
      `If LCase(r.Value) Like "*k" Then`
      `If LCase(r.Value) Like "king*" Then`
      `If LCase(r.Value) Like "king *" Then`
      `If LCase(r.Value) Like "*king*" Then`
     
      `If Not LCase(r.Value) = "*twilight*" Then`
     
      `If LCase(r.Value) Like "*king*" And Not LCase(r.Value) = "*twilight*" Then`
     
      `If LCase(r.Value) Like "* ?" Then`
      `If LCase(r.Value) Like "* ???" Then`
     
      `If LCase(r.Value) Like "* #" Then`
     
      `If LCase(r.Value) Like "*[?]" Then`
     
      `If LCase(r.Value) Like "a*" Or LCase(r.Value) Like "m*" Or LCase(r.Value) Like "g*" Then
      If LCase(r.Value) Like "[amg]*" Then`
     
      `If LCase(r.Value) Like "[j-m]*" Then`
     
      `If Not LCase(r.Value) Like "[j-m]*" Then
      If LCase(r.Value) Like "[!j-m]*" Then`
     
      `If LCase(r.Value) Like "? [!h]*'s ????" Then`