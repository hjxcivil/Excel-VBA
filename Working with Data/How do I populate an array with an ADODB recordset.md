# How do I populate an array with an ADODB recordset

[TOC]

## populate an array

- The Basic

  Sub RecordsetToArray()

      Dim cn As ADODB.Connection:Set cn = New ADODB.Connection
      cn.ConnectionString = _
          "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & ThisWorkbook.Path & "\Answers\Movies.xlsx;" & _
          "Extended Properties='Excel 12.0 Xml; HDR=YES';"
      cn.Open
      
      Dim rs As ADODB.Recordset:Set rs = New ADODB.Recordset:rs.ActiveConnection = cn
      
      rs.Source = "SELECT * FROM [Sheet1$]"
      rs.Open
      	...
      rs.Close
      cn.Close

  End Sub

- The GetRows Method  

  > Dim a As Variant
  >         a = rs.GetRows(10)	= > Variant(0 to 13, 0 to 1199)

- Limiting the Rows Returned

  > a = rs.GetRows(10)	= > = > Variant(0 to 13, 0 to 9)

- Specifiying the Columns Returned

  > a = rs.GetRows(Rows:=10, Fields:=Array("Title", "Release Date", "Run Time")) 
  >
  >  - > Variant(0 to 2, 0 to 9)
  
- Specifying Column Names in the Select Statement

  > rs.Source = "SELECT [Title], [Release Date], [Run Time] FROM [Sheet1$]"
  > a = rs.GetRows()	= > Variant(0 to 2, 0 to 1199)

- Selecting the Top 10 Rows

  > rs.Source = "SELECT TOP 10 [Title], [Release Date], [Run Time] FROM [Sheet1$]"
  > a = rs.GetRows()	= > Variant(0 to 2, 0 to 9)   

- Moving Back to the First Record

  - GetRows moves the cursor to the end of recordset

    > a = rs.GetRows : sh.Range("A1").CopyFromRecordset rs	- - > Empty

  - Moving Back

    > a = rs.GetRows : rs.MoveFirst : sh.Range("A1").CopyFromRecordset rs
