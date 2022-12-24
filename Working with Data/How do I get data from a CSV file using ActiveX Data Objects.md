# How do I get data from a CSV file using ActiveX Data Objects

[TOC]

## Opening a CSV File as a Workbook

> Workbooks.Open DataAnswersPath & "\" & curWs & "\Movies 2011.csv"



## The ActiveX Data Objects Library

`Microsoft ActiveX Data Objects 6.1 Library`

Sub *GetDataFromCSVFiles*()

```
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
...
```

## Creating a Connection String

- ODBC *Standard*:	](https://www.connectionstrings.com/microsoft-text-odbc-driver/)

  > Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=c:\txtFilesFolder\;Extensions=asc,csv,tab,txt;



- Creating the Connection String:

  > Dim txtFilesFolder As String
  >     txtFilesFolder = DataAnswersPath & curWs
  >     cn.ConnectionString = _
  >         "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
  >         "Dbq=" & txtFilesFolder & "\;" & _
  >         "Extensions=asc,csv,tab,txt;

- Selecting Data from a CSV file	

  ```
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.ActiveConnection = cn
  rs.Source = "SELECT * FROM [Movies 2011.csv]"
  rs.Open
  ...
  rs.Close
  ```

- Copying Data from a Recordset 

  > sh.Cells.Clear
  >
  > sh.Range("A2").CopyFromRecordset rs
  >
  > sh.Range("A2").CurrentRegion.EntireColumn.AutoFit

- Basic Union Select Queries

  > rs.Source = "SELECT * FROM [Movies 2011.csv] UNION SELECT * FROM [Movies 2012.csv]"

- Looping Through CSV Files in a Folder

  - get the sqlstring

    > Private Function getSQLString(frompath As String) As String
    >     Dim f As String, SQLString As String
    >     f = Dir(frompath & "\*.csv")
    >     Do Until f = ""
    >         SQLString = SQLString & " UNION SELECT * FROM [" & f & "]"
    >         f = Dir
    >     Loop
    >     SQLString = Mid(SQLString, Len(" UNION ") + 1)
    >     Debug.Print SQLString
    >     getSQLString = SQLString
    > End Function

  - change the source

    > rs.source = getSQLString(txtFilesFolder )

    > SELECT * FROM [Movies 2011.csv] UNION SELECT * FROM [Movies 2012.csv] UNION SELECT * FROM [Movies 2013.csv] UNION SELECT * FROM [Movies 2014.csv] UNION SELECT * FROM [Movies 2015.csv] UNION SELECT * FROM [Movies 2016.csv]


-  Getting the Column Headings

  > Dim i As Integer
  >  For i = 0 To rs.Fields.Count - 1
  >       sh.Cells(1, i + 1).Value = rs.Fields(i).Name
  >  Next i

-  Byte Order Mark Characters: !!!

  >  First Cell may Changed:  FilmID - > 锘縁ilmID

- Sorting the Query Results

  > rs.Source = getSQLString(txtFilesFolder) & " ORDER BY [Run Time] DESC"       

- Adding Criteria to the Query

  > IN Function getSQLString:
  >
  > Do Until f = ""
  >    SQLString = SQLString & " UNION SELECT * FROM [" & f & "] WHERE [Oscar Nominations] > 0" 
  > f = Dir
  >
  > Loop

## Private Const Code 

> Private Const DataAnswersPath As String = "C:\Users\13198\Desktop\Working  with data\Answers\"
> Private Const curWs As String = "vba-adodb-csv"
