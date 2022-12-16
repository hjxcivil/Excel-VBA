# Part 58.2 - SQL for Excel Files - Sorting Rows in a Query

[TOC]

## Sorting Rows in a Query

- The Order By Clause

- Sorting by a Single Column using the Column Name

  > "SELECT * FROM [Film$] ORDER BY [Title]"

- Sorting by Multiple Columns

  > "SELECT * FROM [Film$] ORDER BY [Genre], [Title], [Release Date]"

- Ascending and Descending

  > "SELECT * FROM [Film$] ORDER BY [Oscar Wins] DESC" 
  >
  > "SELECT * FROM [Film$] ORDER BY [Oscar Wins] DESC, [Release Date] ASC"

- Sorting Without Column Headers 'hdr=no

  > "SELECT [F2] AS [Film Name], [F3] AS [Date], [F14] AS [Oscars] FROM [FilmNoHeaders$] ORDER BY [F14] DESC, [F3] ASC"

- Aliases Cannot be used in the Order By Clause with the ACE provider   BELOW IS ERROR

  > "SELECT [F14] AS [Oscars] FROM [FilmNoHeaders$] ORDER BY [Oscars] DESC"

- Using the Column Index of the Select List

  > "SELECT [F2] AS [Film Name], [F3] AS [Date], [F14] AS [Oscars] " & _
  > "FROM [FilmNoHeaders$] " & _
  > "ORDER BY 3 DESC, 2 ASC"    

- Sorting by Hidden Columns

  > "SELECT [Title] " & _
  >         "FROM [Film$] " & _
  >         "ORDER BY [Oscar Wins] DESC, [Release Date] ASC"

