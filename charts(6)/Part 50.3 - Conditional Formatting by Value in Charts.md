# Conditional Formatting in Charts

[TOC]

## Create a Basic Chart in VBA

Sub *FormatColumns*()
    Dim c As Chart
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered

End Sub

## Looping Through the Points Collection

Sub *FormatColumns*()    
     
    Dim c As Chart
    Dim s As Series
    Dim p As Point
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered
    
    Set s = c.SeriesCollection(1)
    
    For Each p In s.Points
        ...
    Next p

End Sub

## Capture the Values Array

Sub *FormatColumns*()
    
    Dim c As Chart
    Dim s As Series
    Dim p As Point
    Dim ValuesArray
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered
    
    Set s = c.SeriesCollection(1)
    
    ValuesArray = s.Values

End Sub

## Looping Through the Values Array

      For n = LBound(ValuesArray) To UBound(ValuesArray)
          If ValuesArray(n) > 10 Then
              Set p = s.Points(n)
              p.Format.Fill.ForeColor.RGB = rgbLime
          End If
      Next n
## Looping Through Multiple Series

Sub *FormatColumns*()

    Dim c As Chart
    Dim s As Series
    Dim p As Point
    Dim ValuesArray
    Dim n As Long
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered
    
    For Each s In c.SeriesCollection
    
        ValuesArray = s.Values
        
        For n = LBound(ValuesArray) To UBound(ValuesArray)
            
            If ValuesArray(n) > 10 Then
                Set p = s.Points(n)
                p.Format.Fill.ForeColor.RGB = rgbLime
            End If
        Next n
    
    Next s

End Sub