### Part 58.3 - SQL for Excel Files - Selecting the Top N Rows

#### Selecting Top Rows from a Query

- Adding the Top Keyword to a Query

- Selecting a Specific Number of Rows

  > "SELECT TOP 5 * FROM [Film$]"

- Using the Order By Clause

  > "SELECT TOP 5 * FROM [Film$] ORDER BY [Run Time] DESC"

- Rows with Tied Values

  > "SELECT TOP 1 * FROM [Film$] ORDER BY [Oscar Wins] DESC"
  
- Removing Tied Results

  - Sorting by a Tie-Breaker Field

    > "SELECT TOP 1 * FROM [Film$] ORDER BY [Oscar Wins] DESC, [Release Date] DESC"

  - Using Unique identifiers as Tie-Breakers

    > "SELECT TOP 1 * FROM [Film$] ORDER BY [Oscar Wins] DESC, [Film ID] DESC"

- Returning a Percentage of Rows

  > "SELECT TOP 1 PERCENT * FROM [Film$]"
  >
  > "SELECT TOP 1 PERCENT  * FROM [Film$] ORDER BY [Oscar Wins] DESC, [Film ID] DESC"

- Rounding Up Rows Returned with Top N Percent ' 10 Results Total

  > "SELECT TOP 1 PERCENT * FROM [Films2019]"  = > 1 result
  >
  > "SELECT TOP 11 PERCENT * FROM [Films2019]" = > 2 results
