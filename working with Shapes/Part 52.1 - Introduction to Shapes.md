# Part 52.1 - Introduction to Shapes

[TOC]

## Working with Shapes in VBA

- The Shapes Collection

  - Shapes
  - Pictures
  - SmartArt
  - Form Controls
  - Chart

- Referring to Shapes

  - Shapes Collection

    > MsgBox Sheet1.Shapes.Count

  - Single Shape

    > Sheet1.Shapes(1).Select
    >
    > Sheet1.Shapes("Heart 4").Select

  - A Range of Shapes

    > Sheet1.Shapes.Range(Array(1, 2, 3)).Select
    >
    > Sheet1.Shapes.Range(Array("Smiley Face 3", "Heart 4", "Picture 5")).Select

  - Selected Shapes

    > Selection.ShapeRange.Fill.ForeColor.RGB = rgbPapayaWhip

  - Using Shape Variable

    > Dim sh As Shape
    >     Set sh = Sheet1.Shapes("Heart 4")
    >     sh.Fill.ForeColor.RGB = rgbHotPink

  - Loop over : for each & for i

- Controlling Shape Size and Position

  - Absolute

    > sh.Left =  100

  - Relative

    > sh.Top = Sheet1.Shapes(3).Top
    >
    > sh.Top = Range("B2").Top
    >
    > sh.Width = Range("B2:C5").Width
    >
    > sh.Height = Range("B2:C5").Height

- Adding Basic Shapes

  > Set sh = Sheet3.Shapes.AddShape(msoShapeHeart, 20, 20, 72, 72)
  > sh.Fill.ForeColor.RGB = rgbHotPink

- Inserting Pictures

  > Set sh = Sheet3.Shapes.AddPicture2( _
  >         Filename:=Environ("UserProfile") & "\Desktop\logo-2x.png", _
  >         LinkToFile:=msoFalse, _
  >         SaveWithDocument:=msoCTrue, _
  >         Left:=100, Top:=20, Width:=-1, Height:=-1, _
  >         Compress:=msoPictureCompressTrue)
  >
  > sh.LockAspectRatio = msoCTrue
  > sh.Width = 100

- Drawing a Button

  > Set sh = Sheet3.Shapes.AddFormControl(xlButtonControl, 50, 100, 200, 50)
  >
  > sh.OnAction = "DeleteSheet3Shapes"

- Adding and Arranging Multiple Shapes

  Sub ***AddMultipleShapes***()

      Const ShWidth As Integer = 50, ShHeight As Integer = 25
      
      Dim i As Integer, j As Integer, ShLeft As Integer, ShTop As Integer
      For j = 0 To 3
          ShTop = j * ShHeight
          For i = 0 To 4
              ShLeft = i * ShWidth
              Dim sh As Shape
              Set sh = _
                  Sheet2.Shapes.AddShape(msoShapeRectangle, ShLeft, ShTop, ShWidth, ShHeight)
          sh.Fill.ForeColor.RGB = rgbHotPink
          Next i
      Next j

  End Sub
