# Working with Dates in VBA

[TOC]

## Writing and Reading Dates in Code

Sub *DateBasics*()

    Dim dt As Date
    
    dt = #5/2/2017#
    
    Debug.Print dt
    MsgBox dt
    Sheet1.Range("A1").Value = dt

End Sub

## Regional Setting

- Writing Unambiguous Dates

  ?#13/5/2017#
  2017/5/13 

  ?#2 may 2017#
  2017/5/2 

  ?#2/5/2017#

  2017/2/5 

## The Date Data Type

- The Range of Dates Available

  Sub *DateRanges*()

      Dim dMin As Date
      Dim dMax As Date
      Dim dMinExcel As Date
      
      dMin = #1/1/100# '100/1/1
      dMax = #12/31/9999# '9999/12/31
      dMinExcel = #1/1/1900# '1900/1/1
      
      MsgBox dMin & vbNewLine & dMax & vbNewLine & dMinExcel
      
      Worksheets.Add
      
      Range("A1").Value = dMin 'Run-time error
      Range("A2").Value = dMax '9999/12/31
      Range("A3").Value = dMinExcel '1900/1/2

  End Sub

- Excel's Leap Year Bug

  Sub *LeapYearBug*() 

      Dim dt As Date
      
      Worksheets.Add
      
      dt = #3/1/1900#
      Range("A1").Value = dt '1900/3/1
      
      dt = #2/28/1900#
      Range("A2").Value = dt '1900/2/29

  End Sub

- Excel incorrectly assumes that the year 1900 is a leap Year

## The Current Date

- Getting the Current Date

  - In Worksheet : "*=today()*"

  - In VBA: 

    Sub *Currentdate*()
        
        Dim dt As Date
        
        dt = Date
        Debug.Print dt
        Range("A5").Value = dt

    End Sub

- Getting the Current Date and Time

  - In Worksheet : "*=now()*"   '*Can be refreshed by F9 or Changing cell*

  - In VBA: 

    Sub *Currentdate*()
        

        dt = Now 
        Debug.Print dt
        Range("A8").Value = dt 'cann't be recal by using F9 ...

    End Sub

- Writing Date Functions into a Worksheet

  Sub *EnterDateFunctions*()

      Range("A10").Value = "=Today()"
      Range("A11").Value = "=Now()" 'can be recaculate by using F9 ...
      Range("A11").NumberFormat = "dd/mm/yyyy hh:mm:ss"

  End Sub

  

## Formatting Dates

- Using Built-In Date Formats

  Sub *FormattingDates*()

      Dim dt As Date
      
      dt = Date
      
      Worksheets.Add
      
      Range("A1").Value = FormatDateTime(dt, vbGeneralDate) '2019/7/16
      Range("A2").Value = FormatDateTime(dt, vbLongDate) '2019年7月16日
      Range("A3").Value = FormatDateTime(dt, vbShortDate) '2019/7/16

  End Sub

- Creating Custom Date Formats

  `Range("A4").Value = Format(dt, "dddd d mmmm yyyy")` '*Tuesday 16 July 2019*
  `Range("A5").Value = Format(dt, "dd/mm/yyyy")` '*16/07/2019*

- The Return Type of The Format Function 

  - *May not always be a Date* ' [a4] above is treated as string not a date

## Calculations with Dates

- Splitting and Reconstructing Dates

  Sub *DateParts*() 

      Dim dt As Date
      
      dt = Date
      
      Worksheets.Add
      
      Range("A1").Value = Year(dt)
      Range("A2").Value = Month(dt)
      Range("A3").Value = Day(dt)
      
      Range("A4").Value = DateSerial(Range("A1").Value, Range("A2").Value, Range("A3").Value)

  End Sub

## Date Functions

- Basic Date Calculations

  - Subtracting Dates

    Sub *BasicDateCalculation*() 

        Dim StartDate As Date, EndDate As Date
        
        StartDate = Date 
        EndDate = Sheet1.Range("A1").Value 
        
        Range("A2").Value = EndDate - StartDate
      Range("A3").Value = DateDiff("d", StartDate, EndDate)
        Range("A4").Value = DateDiff("ww", StartDate, EndDate)
      Range("A5").Value = DateDiff("m", StartDate, EndDate) 
  
  End Sub
  
  - Using Excel's Date Functions
  
    
    `Range("A6").Value = WorksheetFunction.NetworkDays(StartDate, EndDate)`

## Calculating Age in Years

- Limitations of *DateDiff* 'Cannot Caculate complete Year

  `Debug.Print DateDiff("yyyy", #1/1/1990#, #2018/5/12#)`    - - > 28

  `Debug.Print DateDiff("yyyy", #8/15/1990#, #2018/5/12#)`    - - > 28
  
- The Solution

  - ***Excel's *DateDif* Function 

    `Debug.Print Evaluate("DateDif(C1,Today(),""Y"")")`

  - Creating a Function to Calculate Age in Years

    Function *AgeInYears*(ByVal *StartDate* As Date, Optional ByVal *EndDate* As Date) As Integer

        Dim YearsDiff As Integer
        Dim Anniversary As Date
        
        If EndDate = 0 Then EndDate = Date
        
        YearsDiff = DateDiff("yyyy", StartDate, EndDate)
        Anniversary = DateAdd("yyyy", YearsDiff, StartDate)
        
        If Anniversary > EndDate Then YearsDiff = YearsDiff - 1
        
        AgeInYears = YearsDiff

    End Function
