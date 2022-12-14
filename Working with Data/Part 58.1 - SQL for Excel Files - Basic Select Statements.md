# Part 58.1 - SQL for Excel Files - Basic Select Statements

[TOC]

## Writing Basic SQL Queries for Excel

- Connect to an Excel File

- Getting all Columns from a:

  - Worksheet 	  *"[Film$]"*
  - Specific Range of Cells   *"[FilmYears$A2:D12]"*
  - Named Range    *"[Films2019]"*
  - Worksheet Scoped Range Names    "[FilmYears2$Films2019Local]"
  - Tables Without a Header Row  "[FilmNoHeaders$]" 'HDR=NO'

- SQL Select Statements

  > "SELECT * FROM [Film$]"
  >
  > "SELECT * FROM [Film$A1:E11]"
  >
  > "SELECT * FROM [Films2019]"
  >
  > "SELECT * FROM [FilmYears2$Films2019Local]"

- Selecting Specific Columns

  - by Name

    > "SELECT [Title], [Run Time], [Release Date], [Oscar Wins] FROM [Film$]"

  - Without Names    'HDR=NO

    > "SELECT [F2], [F4], [F3], [F14] FROM [FilmNoHeaders$]"
    >
    > "SELECT [F2] AS [Title], [F4] AS [Minutes], [F3] AS [Release Date], 
    >
    > [F14] AS [Oscar Wins] FROM [FilmNoHeaders$]"

- Using Column and Table Aliases

  > "SELECT [f].[Title] AS [Film Name], [f].[Run Time], [f].[Release Date], [f].[Oscar Wins] FROM [Film$] AS [f]"  

- Code Layout  http://poorsql.com/

```
SELECT [f].[Title] AS [Film Name]
	,[f].[Run Time]
	,[f].[Release Date]
	,[f].[Oscar Wins]
FROM [Film$] AS [f]
```

