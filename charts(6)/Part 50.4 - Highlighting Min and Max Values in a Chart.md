# Formatting Min and Max Chart Values

[TOC]

## Creating a Basic Column Chart in VBA

Sub *HighlightMinMax*()

    Dim c As Chart
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData ThisWorkbook.Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered

End Sub

## Looping Through Chart Series

Sub *HighlightMinMax*()

    Dim c As Chart
    Dim s As Series
    Dim ValuesArray
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData ThisWorkbook.Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered
    
    For Each s In c.SeriesCollection
        ValuesArray = s.Values
    Next s

End Sub

## Calculating the Min and Max Values

Sub *HighlightMinMax*()

    Dim c As Chart
    Dim s As Series
    Dim ValuesArray
    Dim MaxVal As Double, MinVal As Double
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData ThisWorkbook.Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered
    
    For Each s In c.SeriesCollection
    
        ValuesArray = s.Values
        
        MaxVal = WorksheetFunction.Max(ValuesArray)
        MinVal = WorksheetFunction.Min(ValuesArray)
        
        Debug.Print MaxVal, MinVal
        
    Next s

End Sub

## Looping Through the Values Array

Sub *HighlightMinMax*()

    Dim c As Chart
    Dim s As Series
    Dim ValuesArray
    Dim MaxVal As Double, MinVal As Double
    Dim n As Long
    Dim p As Point
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData ThisWorkbook.Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlColumnClustered
    
    For Each s In c.SeriesCollection
    
        ValuesArray = s.Values
        
        MaxVal = WorksheetFunction.Max(ValuesArray)
        MinVal = WorksheetFunction.Min(ValuesArray)
        
        For n = LBound(ValuesArray) To UBound(ValuesArray)
            
            Set p = s.Points(n)
            ... ...
            
        Next n
        
    Next s

End Sub

## Applying Formatting to Points

```
If ValuesArray(n) = MaxVal Then
	p.Format.Fill.ForeColor.RGB = rgbLime
ElseIf ValuesArray(n) = MinVal Then
    p.Format.Fill.ForeColor.RGB = rgbRed
End If
```

## Line Charts and Markers

Sub *HighlightMinMax*()

    Dim c As Chart
    Dim s As Series
    Dim ValuesArray
    Dim MaxVal As Double, MinVal As Double
    Dim n As Long
    Dim p As Point
    
    Set c = ThisWorkbook.Charts.Add
    
    c.SetSourceData ThisWorkbook.Worksheets("Sheet1").Range("A1").CurrentRegion
    c.ChartType = xlLine
    
    For Each s In c.SeriesCollection
    
        s.MarkerStyle = xlMarkerStyleCircle
        ValuesArray = s.Values
        
        MaxVal = WorksheetFunction.Max(ValuesArray)
        MinVal = WorksheetFunction.Min(ValuesArray)
    
        For n = LBound(ValuesArray) To UBound(ValuesArray)
        
            Set p = s.Points(n)
            
            If ValuesArray(n) = MaxVal Then
                p.MarkerForegroundColor = rgbLime
                p.MarkerBackgroundColor = rgbLime
            ElseIf ValuesArray(n) = MinVal Then
                p.MarkerForegroundColor = rgbRed
                p.MarkerBackgroundColor = rgbRed
            End If
        
        Next n
        
    Next s

End Sub