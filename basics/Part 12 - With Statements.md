# Using With Statements

[TOC]

## Simple With Statements

Sub *FormatFilmReleaseDates*()

    With Range("C3:C15")
        .Interior.Color = rgbAquamarine
        .Font.Color = rgbRed
        .Font.Size = 12
        .NumberFormat = "dddd dd mmm yyyy"
    End With

End Sub

## Using Complex References

Sub *FormatFilmReleaseDates*()

    With Worksheets("Sheet1").Range("C3", Worksheets("Sheet1").Range("C2").End(xlDown))
        .Interior.Color = rgbAquamarine
        .Font.Color = rgbRed
        .Font.Size = 12
        .NumberFormat = "dddd dd mmm yyyy"
    End With

End Sub



## Referencing Other Objects

Sub *FormatFilmReleaseDates*()

    With Worksheets("Sheet1").Range("C3", Worksheets("Sheet1").Range("C2").End(xlDown))
        .Interior.Color = rgbAquamarine
        
        Worksheets("Sheet3").Cells.Interior.Color = .Interior.Color
        
        .Font.Color = rgbRed
        .Font.Size = 12
        .NumberFormat = "dddd dd mmm yyyy"
    End With

End Sub