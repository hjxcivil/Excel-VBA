- The Basics of Writing VBA Code
  - Beginning a Subroutine
  - Laying Out Code Neatly
  - Writing Comments: ' or REM
- Writing VBA Instructions
  - Basic VBA Grammar
  - Changing the Value of Cells
  - Formatting Cells
- Running VBA Code
  - Saving Files Containing Code
  - Running a Subroutine
  - Reopening Files and Security

## The First VBA Macro

Sub *CreateAndLabelNewSheet*()

    'Create a new worksheet
    'Object.Method
    Worksheets.Add
    
    'add titles to cells
    'Object.Property = Value
    Range("A1").Value = "Created by"
    Range("A2").Value = "Created on"
    Range("A3").Value = "Version"
    
    'add user values to cells
    Range("B1").Value = Environ("UserName")
    Range("B2").Value = Date
    Range("B3").Value = 1
    
    'format titles
    Range("A1:A3").Font.Color = vbBlue
    Range("A1:A3").Interior.Color = rgbPaleTurquoise

End Sub