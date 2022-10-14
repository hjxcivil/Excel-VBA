# Using Data Points on Scatter Charts

[TOC]

## The Problem with Data Labels in Scatter Charts

- Version2013 or Later
  - Add Data Labels : Right Click Data Points Choose Add Data Labels Add Data Labels again Default display Y-axis Value
  - Format Data Labels: Label Options contains *Value from Cells*
- Version2010: 
  - Label Options is Shortly , has no *Value from Cells*

## Using a For Each Loop to Label Each Point

- The Basic Solution to Add Data Labels

  - Creating a Basic Subroutine

    Sub *CreateDataLabels*()

        ..

    End Sub

  - Declaring Variables

    `Dim r As Range`
    `Dim FilmNames As Range`
    `Dim FilmCounter As Integer`
    `Dim FilmSeries As Series`

  - Set a Reference to a Range of Cells

    `Sheet1.Select`
    `Set FilmNames = Range("A2", Range("A1").End(xlDown))`

  - Set a Reference to a Data Series

    `Set FilmSeries = Sheet1.ChartObjects(1).Chart.SeriesCollection(1)`  *OR*

    `Set FilmSeries = Chart1.SeriesCollection(1)`

  - Referring to Chart Sheets : See upper

  - Enabling Data Labels for a Series

    `FilmSeries.HasDataLabels = True`

  - Creating a Basic For Each Loop

  - Adding Labels to a Data Point

  - Running the Code

    Sub *CreateDataLabels*()

        Dim r As Range
        Dim FilmNames As Range
        Dim FilmCounter As Integer
        Dim FilmSeries As Series
        
        Sheet1.Select
        Set FilmNames = Range("A2", Range("A1").End(xlDown))
        Set FilmSeries = Sheet1.ChartObjects(1).Chart.SeriesCollection(1)
        
        FilmSeries.HasDataLabels = True
        
        For Each r In FilmNames
        
            FilmCounter = FilmCounter + 1
            FilmSeries.Points(FilmCounter).DataLabel.Text = r.Value
        
        Next r

    End Sub

## Labelling Multiple Charts

![2](C:\Users\13198\Documents\2.jpg)

Sub *CreateDataLabelsForMultipleCharts*()

    Dim r As Range
    Dim FilmNames As Range
    Dim FilmCounter As Integer
    Dim FilmSeries As Series
    Dim ch As ChartObject
    
    Sheet2.Select
    Set FilmNames = Range("A2", Range("A1").End(xlDown))
    
    For Each ch In Sheet2.ChartObjects
    
        Set FilmSeries = ch.Chart.SeriesCollection(1)
        FilmSeries.HasDataLabels = True
        FilmCounter = 0
        
        For Each r In FilmNames
    
            FilmCounter = FilmCounter + 1
            FilmSeries.Points(FilmCounter).DataLabel.Text = r.Value
    
        Next r
    Next ch

End Sub

![3](C:\Users\13198\Documents\3.jpg)

## Labelling Multiple Series

Sub *CreateDataLabelsForMultipleSeries*()

    Dim r As Range
    Dim FilmNames As Range
    Dim FilmCounter As Integer
    Dim FilmSeries As Series
    
    Sheet3.Select
    Set FilmNames = Range("A2", Range("A1").End(xlDown))
    
    For Each FilmSeries In Sheet3.ChartObjects(1).Chart.SeriesCollection
        FilmSeries.HasDataLabels = True
    Next FilmSeries
    
    For Each r In FilmNames
    
        FilmCounter = FilmCounter + 1
        
        For Each FilmSeries In Sheet3.ChartObjects(1).Chart.SeriesCollection
            FilmSeries.Points(FilmCounter).DataLabel.Text = r.Value
        Next FilmSeries
    
    Next r

End Sub

## Labelling Multiple Chart Sheets

Sub *CreateDataLabelsForMultipleChartSheets*()

    Dim r As Range
    Dim FilmNames As Range
    Dim FilmCounter As Integer
    Dim FilmSeries As Series
    Dim ch As Chart
    
    Sheet2.Select
    Set FilmNames = Range("A2", Range("A1").End(xlDown))
    
    For Each ch In Charts
    
        Set FilmSeries = ch.SeriesCollection(1)
        FilmSeries.HasDataLabels = True
        FilmCounter = 0
        
        For Each r In FilmNames
    
            FilmCounter = FilmCounter + 1
            FilmSeries.Points(FilmCounter).DataLabel.Text = r.Value
    
        Next r
    Next ch

End Sub