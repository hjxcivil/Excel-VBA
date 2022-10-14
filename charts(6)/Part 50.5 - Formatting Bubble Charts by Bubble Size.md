# Formatting Bubble Charts

[TOC]

## Creating a Bubble Chart in VBA

Sub *FormatBubbles*()

    Dim c As Chart
    
    Set c = Worksheets("Sheet1").Shapes.AddChart2(XlChartType:=xlBubble).Chart
    
    c.SetSourceData Worksheets("Sheet1").Range("A2", Worksheets("Sheet1").Range("A2").End(xlDown).End(xlToRight))

End Sub

- Moving the Chart to a Separate Sheet

  `c.Location xlLocationAsNewSheet`

- Referring to a Series in the Chart

  `Set s = c.SeriesCollection(1)`

- Referring to the Bubble Sizes

  `Debug.Print s.BubbleSizes` -> "Sheet1!$C$2:$C$11"

- Assigning Bubble Sizes to an Array

  `BubbleSizeArray = Range(s.BubbleSizes)`

- Finding the Min and Max Bubble Sizes

  `MaxSize = WorksheetFunction.Max(Range(s.BubbleSizes))`
  `MinSize = WorksheetFunction.Min(Range(s.BubbleSizes))`

- Looping Through the Array and Formatting a Bubble

      For n = LBound(BubbleSizeArray, 1) To UBound(BubbleSizeArray, 1)
      
      Set p = s.Points(n)
          
          If BubbleSizeArray(n, 1) > MaxSize Then
              p.Format.Fill.ForeColor.RGB = rgbLime
          ElseIf BubbleSizeArray(n, 1) = MinSize Then
              p.Format.Fill.ForeColor.RGB = rgbRed
          End If
      
      Next n

## Completed Subroutine

Sub *FormatBubbles*()

    Dim c As Chart
    Dim s As Series
    Dim BubbleSizeArray
    Dim MaxSize As Double, MinSize As Double
    Dim n As Long
    Dim p As Point
    
    Set c = Worksheets("Sheet1").Shapes.AddChart2(XlChartType:=xlBubble).Chart
    
    c.SetSourceData Worksheets("Sheet1").Range("A2", Worksheets("Sheet1").Range("A2").End(xlDown).End(xlToRight))
    
    Set s = c.SeriesCollection(1)
    
    BubbleSizeArray = Range(s.BubbleSizes)
    MaxSize = 10 'WorksheetFunction.Max(Range(s.BubbleSizes))
    MinSize = WorksheetFunction.Min(Range(s.BubbleSizes))
    
    For n = LBound(BubbleSizeArray, 1) To UBound(BubbleSizeArray, 1)
    
        Set p = s.Points(n)
        
        If BubbleSizeArray(n, 1) > MaxSize Then
            p.Format.Fill.ForeColor.RGB = rgbLime
        ElseIf BubbleSizeArray(n, 1) = MinSize Then
            p.Format.Fill.ForeColor.RGB = rgbRed
        End If
    
    Next n
    
    c.Location xlLocationAsNewSheet

End Sub