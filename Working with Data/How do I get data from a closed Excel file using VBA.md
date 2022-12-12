# How do I get data from a closed Excel file using VBA

[TOC]

## Referencing the Library

`Microsoft ActiveX Data Objects 6.1 Library`

Sub *GetDataFromClosedFile*()

    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    ...

## Connecting to a Closed Workbook

> cn.ConnectionString = _
>         "Provider=Microsoft.ACE.OLEDB.12.0;" & _
>         "Data Source=" & DataAnswersPath & "\Movies.xlsx;" & _
>         "Extended Properties='Excel 12.0 Xml; HDR=YES';"
> cn.Open
>
> cn.Close

- Selecting Data into a Recordset	

  ```
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.ActiveConnection = cn
  rs.Source = "SELECT * FROM [Sheet1$]"
  rs.Open
  ...
  rs.Close
  ```

- Copying Data from a Recordset 

  > sh.Range("A1").CurrentRegion.Offset(1, 0).Clear
  >
  > sh.Range("A2").CopyFromRecordset rs
  >
  > sh.Range("A1").CurrentRegion.EntireColumn.AutoFit



- Copying a Specific Range of Cells

  > rs.Source = "SELECT * FROM [Sheet1$A1:D11]"

-  Returning Specific Columns

​		`rs.Source = "SELECT [Run Time], [Studio], [Budget]  FROM [Sheet1$]"`

-  Returning Columns Without Column Headers 

​		`rs.Source = "SELECT [F4], [F7], [F11]  FROM [Sheet1$]" 'HDR=NO`

- Combining Criteria and Wildcards

​	 	rs.Source = _
​                "SELECT * FROM [Sheet1$] " _
​                & "WHERE [Oscar Wins] >=1 AND [Genre] = 'Science Fiction' AND [Title] LIKE 'star%'"

## The Whole Code 

Option Explicit

> Private Const DataAnswersPath As String = "C:\Users\13198\Desktop\Working  with data\Answers"
> Private Const curWs As String = "copy-from-closed-workbook"
> Private Enum SqlMethod
>     mAll = 0
>     mSpecificRange = 1
>     mSpecificColumn = 2
>     mNoHeadings = 3
>     mCriteria = 4
> End Enum

Sub GetDataFromClosedFile()

```
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection

cn.ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=" & DataAnswersPath & "\Movies.xlsx;" & _
    "Extended Properties='Excel 12.0 Xml; HDR=YES';"

cn.Open

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.ActiveConnection = cn

Dim Method As SqlMethod
Method = mAll
Select Case Method
    Case 0: rs.Source = "SELECT * FROM [Sheet1$]"
    Case 1: rs.Source = "SELECT * FROM [Sheet1$A1:D11]"
    Case 2: rs.Source = "SELECT [Run Time], [Studio], [Budget]  FROM [Sheet1$]"
    Case 3: rs.Source = "SELECT [F4], [F7], [F11]  FROM [Sheet1$]" 'HDR=NO
    Case 4: rs.Source = _
            "SELECT * FROM [Sheet1$] " _
            & "WHERE [Oscar Wins] >=1 AND [Genre] = 'Science Fiction' AND [Title] LIKE 'star%'"
            
End Select
  
rs.Open

    Dim sh As Worksheet
    Set sh = Worksheets(curWs)
    
    sh.Range("A1").CurrentRegion.Offset(1, 0).Clear
    
    sh.Range("A2").CopyFromRecordset rs
    sh.Range("A1").CurrentRegion.EntireColumn.AutoFit

rs.Close
cn.Close
```

End Sub
