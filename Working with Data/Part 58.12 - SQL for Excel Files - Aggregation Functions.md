# Part 58.12 - SQL for Excel Files - Aggregation Functions

[TOC]

## Aggregation Functions in SQL

- Summing a Column

  > "SELECT Sum([f].[Run Time]) AS [Total Run Time] FROM [Film$] AS [f]"

- Mixing Aggregated and Non-Aggregated Columns

  > "SELECT Sum([f].[Run Time]) AS [Total Run Time], f.[Title] FROM [Film$] AS [f]"    - > ERROR!

- Counting Values and Rows

  > "SELECT Sum([f].[Run Time]) AS [Total Run Time], Count(f.[Title]) AS [Count of Titles] FROM [Film$] AS [f]"
  >
  > Count(*) AS [Count of Rows]

- Min and Max Values

  > Max([f].[Run Time])
  >
  > Min(f.[Title])

- Caculating Averages

  > Avg([f].[Run Time])
  >
  > Round(Avg([f].[Run Time]), 2)

- Standard Deviation and Variance

  > StDev StDevP Var VarP

- Where clause apply before Aggregation get calculated

- Aggregates and Other Calculations

  - Using Aggregates in Other Calculations

    > "SELECT Max(f.[Run Time]) - Min(f.[Run Time]) AS [Run Time Range]       

  - Aggregating a Calculated Value

    > "SELECT Avg(f.[Box Office] - f.[Budget]) AS [Average Profit]

- Dealing with Nulls

  - Aggregating ignore the Nulls thus Count(f.[Title]) may > Count(f.[Budget])

  - Desc

    > "SELECT " & _
    >             "Count(**) AS [Count of Rows]" & _
    >             ",Count(f.[Budget]) AS [Count of Budgets]" & _
    >             ",Avg(f.[Budget]) AS [Average of Budgets]" & _
    >             ",Sum(f.[Budget]) / Count(f.[Budget]) AS [Average of Budgets 2]" & _
    >             ",Sum(f.[Budget]) / Count(*) AS [Average of Budgets 3]"  & _
    >
    > " FROM [Film$] AS [f]"

  - Replacing Nulls With a value

    > Avg(IIf(IsNull(f.[Budget]), 0, f.[Budget])) AS [Average of Budgets 4]

  - Result

    | Count of  Rows | Count of Budgets | Average of Budgets | Average of Budgets 2 | Average of Budgets 3 | Average of Budgets 4 |
    | -------------- | ---------------- | ------------------ | -------------------- | -------------------- | -------------------- |
    | 1200           | 1076             | 54999235.76        | 54999235.76          | 49315981.4           | 49315981.4           |

  - Filtering Out Nulls

    > WHERE " & _
    >             "f.[Budget] IS NOT NULL AND f.[Box Office] IS NOT NULL"

    | Count of  Rows | Count of Budgets | Average of Budgets | Average of Budgets 2 | Average of Budgets 3 | Average of Budgets 4 |
    | -------------- | ---------------- | ------------------ | -------------------- | -------------------- | -------------------- |
    | 1046           | 1046             | 56265933.72        | 56265933.72          | 56265933.72          | 56265933.72          |
