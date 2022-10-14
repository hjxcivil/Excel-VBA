# Crating, Editing and Formatting Charts

[TOC]

## Creating a New Chart

- Creating a Chart with a Single Cell Selected

  `Charts.Add`

- Ensuring a Valid Range is Selected

  ```
  wsAwards15.Select
  Range("A3").Select
  Charts.Add
  ```

- Deleting Every Chart from a Workbook

  ```
  Application.DisplayAlerts = False
  On Error Resume Next
  Charts.Delete
  ```

## Selecting Data for the Chart 

- Selecting a Block of Cells

  ```
  wsAwards16.Select
  Range("A3:C8").Select
  Charts.Add
  ```

- Selecting Non-Contiguous Columns

  ```
  wsAwards16.Select
  Range("A3:A24, C3:C24").Select
  Charts.Add
  ```

- Selecting Non-Contiguous Rows

  ```
  wsAwards15.Select
  Range("A3:C3, A8:C8 , A11:C11, A16:C16").Select
  Charts.Add
  ```

- Selecting to the End of a List

  ```
  wsAwards15.Select
  Range("A3", Range("B3").End(xlDown)).Select
  Charts.Add
  ```

- The Union Method

      wsAwards16.Select
      Union( _
          Range("A3", Range("A3").End(xlDown)), _
          Range("C3", Range("C3").End(xlDown))).Select
          
      Charts.Add

- Declaring Variabls to Loop Over a Range

  Sub *CreateOscarWinnerChart*()

      Dim ChartCells As Range
      Dim Film As Range
      Dim Films As Range
      
      wsAwards16.Select
      
      Set Films = Range("A4", Range("A3").End(xlDown))
      Set ChartCells = Range("A3", Range("A3").End(xlToRight))
      
      For Each Film In Films
          If Film.Offset(0, 2).Value > 0 Then
              Set ChartCells = Union(ChartCells, Range(Film, Film.End(xlToRight)))
          End If
      Next Film
      
      ChartCells.Select
      Charts.Add

  End Sub

- Placing the Chart Before another Sheet

  Sub *CreatingChartsAndChoosingPositions*()

      wsAwards15.Select
      Range("A3:C19").Select
      
      Charts.Add wsMenu 'Blank Chart with No Data 

  End Sub

## Changing the Source Data and Chart Type

- Changing the Source Data after Creating the Chart

  `Charts.Add`
  `ActiveChart.SetSourceData wsAwards15.Range("A3:C19")`
  
- Creating and the Moving a Chart

      wsAwards15.Select
      Range("A3:C19").Select
      
      Charts.Add
      ActiveChart.Move wsMenu

- Creating a Chart as the First or Last Sheet

      wsAwards15.Select
      Range("A3:C19").Select
      
      Charts.Add
      ActiveChart.Move After:=Sheets(Sheets.Count)'ActiveChart.Move Sheets(1)

- Referring to Existing and New Charts

  Sub *ReferencingCharts*()

      Dim ch As Chart
      
      wsAwards15.Select
      Range("A3:C19").Select
      
      Set ch = Charts.Add
      
      ch.SetSourceData ...

  End Sub

- Using the Chart Wizard to Edit a Chart

  Sub *EditingCharts*()

      Dim ch As Chart
      
      wsAwards15.Select
      Range("A3:C19").Select
      
      Set ch = Charts.Add
      
      ch.ChartWizard Source:=wsAwards16.Range("A3:C10")

  End Sub

- Changing the Chart Type

  Sub *EditingCharts*()

      Dim ch As Chart
      
      wsAwards15.Select
      Range("A3:B19").Select
      
      Set ch = Charts.Add
      ch.ChartType = xl3DPieExploded ' xlColumnClustered xl3DColumnClustered
  End Sub

## Controlling Chart Elements

- Adding a Legend and Title

  Sub *EditingCharts*()

      Dim ch As Chart
      
      wsAwards15.Select
      Range("A3").Select
      
      Set ch = Charts.Add
      
      ch.ChartType = xlColumnClustered
      ch.HasLegend = True
      ch.HasTitle = True
      ch.ChartTitle.Text = "Noms vs. Wins"


  End Sub

- The Add and Add2 Methods

- Adding Axis Title

  Sub *EditingCharts*()

      Dim ch As Chart
      
      wsAwards15.Select
      Range("A3").Select
      
      Set ch = Charts.Add
      
      ch.ChartType = xlColumnClustered
      ch.HasLegend = True
      ch.HasTitle = True
      ch.ChartTitle.Text = "Noms vs. Wins"
      
      ch.Axes(xlCategory).HasTitle = True
      ch.Axes(xlCategory).AxisTitle.Text = "Film Name"
      
      ch.Axes(xlValue).HasTitle = True
      ch.Axes(xlValue).AxisTitle.Text = "Quantity"


  End Sub

- Adding Data Labels

  ```
  ch.SeriesCollection(1).HasDataLabels = True
  ch.SeriesCollection(2).HasDataLabels = True
  ```

  ​	or

      For Each s In ch.SeriesCollection
          s.ApplyDataLabels
      Next s

## Using Chart Layouts, Colours and Styles

- Changing Chart Layout

  Sub *ChangeChartLayout*()

      Dim ch As Chart
      Dim i As Integer
      
      wsAwards15.Select
      Range("A3").Select
      
      For i = 1 To 11
          Set ch = Charts.Add
      
          ch.ApplyLayout i
          
          ch.HasTitle = True
          ch.ChartTitle.Text = "Chart Layout " & i
      Next i

  End Sub

- Changing Chart Colours

  Sub *ChangeChartColours*()

      Dim ch As Chart
      Dim i As Integer
      
      wsAwards15.Select
      Range("A3").Select
      
      For i = 1 To 26
          Set ch = Charts.Add
          ch.ChartColor = i
          ch.HasTitle = True
          ch.ChartTitle.Text = "Chart Color " & i
      Next i

  End Sub

- Using Chart Styles

  Sub *ChangeChartStyles*()

      Dim ch As Chart
      Dim i As Integer
      
      wsAwards15.Select
      Range("A3").Select
      
      For i = 1 To 352
          
          If i = 49 Then i = 201
          
          Set ch = Charts.Add
          ch.ChartStyle = i
          ch.HasTitle = True
          ch.ChartTitle.Text = "Chart Style " & i
      Next i

  End Sub

## Creating and Applying Chart Templates

Sub *CreateNewChartWithTemplate*()

    Dim ch As Chart
    
    wsAwards15.Select
    Range("A3").Select
    
    Set ch = Charts.Add
    
    'ch.ApplyChartTemplate "C:\YourTemplateFolder\YourTemplateName.crtx"

End Sub