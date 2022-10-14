# Working with ChartObjects

[TOC]

## Creating a Chart in a Worksheet

- Using the *AddChart* Method

  Sub *CreateChartobject*()

      wsAwards15.Select
      Range("A3").Select
      
      wsAwards15.Shapes.AddChart

  End Sub

- Using the *AddChart2* Method

  `wsAwards15.Shapes.AddChart2`

- The Add Method of the *ChartObjects* Collection

  `wsAwards15.ChartObjects.Add 300, 30, 600, 300`

## Setting the Source Data of the ChartObject

Sub *CreateChartobject*()

	wsAwards15.Select
	Range("A3").Select
	
	wsAwards15.ChartObjects.Add 300, 30, 600, 300
	wsAwards15.ChartObjects(1).Chart.SetSourceData wsAwards15.Range("A3").CurrentRegion

End Sub	

- Deleting Every ChartObject

  Sub *DeleteAllChartObjects*()

      On Error Resume Next
      ActiveSheet.ChartObjects.Delete

  End Sub

- Specifying the Chart Type, Position and Size

  Sub *CreateChartobject*()

      wsAwards15.Select
      Range("A3").Select
      
      wsAwards15.Shapes.AddChart XlChartType.xl3DColumnClustered, 300, 30, 600, 300

  End Sub

- Choosing a Chart Style with the *AddChart2* Method

  Sub *CreateChartobject*()

      wsAwards15.Select
      
      DeleteAllChartObjects
      
      Range("A3").Select
      
      wsAwards15.Shapes.AddChart2 330, XlChartType.xl3DColumnClustered, 300, 30, 600, 300

  End Sub

- Selecting a Specific Range  

      Range("A3:C11").Select
      
      wsAwards15.Shapes.AddChart2 330, XlChartType.xl3DColumnClustered, 300, 30, 600, 300

- Selecting a Variable Number of Rows 

      Range("A3", Range("B3").End(xlDown)).Select
      
      wsAwards15.Shapes.AddChart2 -1, XlChartType.xl3DColumnClustered, 300, 30, 600, 300

- Selecting Non-Adjacent Columns

      Union( _
              Range("A3", Range("A3").End(xlDown)), _
              Range("C3", Range("C3").End(xlDown))).Select
              
      wsAwards15.Shapes.AddChart2 -1, XlChartType.xl3DColumnClustered, 300, 30, 600, 300

- Selecting Chart Data Conditionally

  Sub *ChartOfRecentFilms*()

      Dim Film As Range
      Dim Films As Range
      Dim RecentFilms As Range
      
      wsMoney.Select
      
      DeleteAllChartObjects
      
      Set Films = Range("B4", Range("B3").End(xlDown))
      Set RecentFilms = Range("B3:C3")
      
      For Each Film In Films
            If Film.Offset(0, 2).Value > 2010 Then
                Set RecentFilms = Union(RecentFilms, Range(Film, Film.Offset(0, 1)))
            End If
            
        Next Film
        
        RecentFilms.Select
        
        ActiveSheet.Shapes.AddChart XlChartType.xl3DColumn, 300, 30, 600, 300

  End Sub

## Referencing the ChartObject and the Chart

- Creating Multiple ChartObjects

  Sub *CreateMultipleCharts*()

      Dim ch2015 As Shape
      Dim ch2016 As ChartObject
      
      wsMenu.Select
      DeleteAllChartObjects
      
      Set ch2015 = wsMenu.Shapes.AddChart(Left:=0, Top:=0, Width:=600, Height:=300)
      Set ch2016 = wsMenu.ChartObjects.Add(Left:=600, Top:=0, Width:=600, Height:=300)
      
      ch2015.Chart.SetSourceData wsAwards15.Range("A3").CurrentRegion
      ch2016.Chart.SetSourceData wsAwards16.Range("A3").CurrentRegion

  End Sub

## Editing and Formatting the Chart

- Changing the Layout and Formatting

  ```
  With ch2015.Chart
      .SetSourceData wsAwards15.Range("A3").CurrentRegion
      .ChartType = xl3DColumnClustered
      .ApplyLayout 2
      .ChartColor = 4
      .HasTitle = True
      .ChartTitle.Text = "Noms vs. Wins"
  End With
  ```
  
- Looping Over ChartObjects

  Sub *LoopOverChartObjects*()

      Dim co As ChartObject
      
      wsMenu.Select
      
      For Each co In wsMenu.ChartObjects
          With co.Chart
              .ChartType = xl3DColumnClustered
              .ApplyLayout 2
              .ChartColor = 4
              .HasTitle = True
              .ChartTitle.Text = "Noms vs. Wins"
          End With
      Next co

  End Sub

- Changing Chart Styles

  Sub *ChangeChartStyles*()

      Dim co As ChartObject
      
      wsMenu.Select
      
      For Each co In wsMenu.ChartObjects
          With co.Chart
              .ChartStyle = 331
              .ChartColor = 3
          End With
      Next co

  End Sub

## Converting ChartObjects to Charts and Back

- Converting ChartObjects to Chart Sheets

  Sub *ChangeChartLocation*()

      Dim co As ChartObject
      
      wsMenu.Select
      
      For Each co In wsMenu.ChartObjects
          With co.Chart
              .Location xlLocationAsNewSheet
          End With
      Next co

  End Sub

- Converting Chart Sheets to ChartObjects

  Sub *ConvertToChartObjects*()

      Dim ch As Chart
      
      For Each ch In ThisWorkbook.Charts
          ch.Location xlLocationAsObject, "Menu"
      Next ch

  End Sub

## Controlling the Size and Position

- Positioning and Sizing a ChartObject Based on a Range

  Sub *PositionChartToCell*()

      Dim co As ChartObject
      Dim Films As Range
      
      Set Films = wsAwards15.Range("A3").CurrentRegion
      Set co = wsAwards15.ChartObjects(1)
      
      co.Left = Range("A3").End(xlToRight).Offset(0, 2).Left
      co.Top = Films.Top
      
      co.Height = Films.Height
      co.Width = Films.Width * 3

  End Sub

- Positioning Multiple Charts

  Sub *PositionMultipleChartObjects*()

      Dim i As Integer
      Dim ChartSizeRange As Range
      Dim ChartsInRow As Integer
      
      ChartsInRow = 2
      Set ChartSizeRange = Range("A1:J15")
      wsMenu.Select
      
      For i = 1 To wsMenu.ChartObjects.Count
      
          With wsMenu.ChartObjects(i)
              .Width = ChartSizeRange.Width
              .Height = ChartSizeRange.Height
              
              .Top = Int((i - 1) / ChartsInRow) * ChartSizeRange.Height
              .Left = ((i - 1) Mod ChartsInRow) * ChartSizeRange.Width
          
          End With
      Next i

  End Sub

- Moving Every ChartObject to One Sheet

  Sub *MoveAllChartObjectsToMenu*()

      Dim ws As Worksheet
      Dim co As ChartObject
      For Each ws In ThisWorkbook.Worksheets
          If Not ws Is wsMenu Then
              For Each co In ws.ChartObjects
                  co.Chart.Location xlLocationAsObject, "Menu"
              Next co
          End If
      Next ws

  End Sub