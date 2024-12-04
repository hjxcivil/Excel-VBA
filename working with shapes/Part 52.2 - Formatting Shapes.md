## Part 52.2 - Formatting Shapes

### Formatting Shapes in VBA

- Set up an example

- Changing Fill Colors

  > s.Fill.ForeColor.RGB = 56342	 '0-16777215 (256^3-1)

- RGB, Theme Colors and Scheme Colors

  - Using RGB Colors

    - vbconstants

      > s.Fill.ForeColor.RGB = vbRed

    - color constants

      >  s.Fill.ForeColor.RGB = ColorConstants.vbMagenta

    - rgb constants

      > s.Fill.ForeColor.RGB = XlRgbColor.rgbMistyRose

    - rgb function [NHS Identity Guidelines | Colours (england.nhs.uk)](https://www.england.nhs.uk/nhsidentity/identity-guidelines/colours/)

      > s.Fill.ForeColor.RGB = RGB(0, 94, 184)
      >
      > ?rgbMistyRose -> 14804223 
      >
      
    - Creating Custom Enumerations(NHS)
    
      > ?RGB(0, 94, 184) -> 12082688 
      >
      > Public Enum NHSColours
      >     NHSBlue = 12082688
      > End Enum
      >
      > s.Fill.ForeColor.RGB = NHSBlue
    
  - Using Theme Colors
  
    - theme Color
  
      > s.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent3
  
    - Changing the Document Theme: Will Change Color and fonts and others...
  
      > ThisWorkbook.ApplyTheme "C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Ion.thmx"
      >
      > ThisWorkbook.ApplyTheme "C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Office Theme.thmx"
  
    - Changing Theme Colors
  
      > ThisWorkbook.Theme.ThemeColorScheme.Load _
      >         "C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Colors\Red.xml"
  
  - Using Scheme Colors
  
    - Changing Scheme Colors
  
      > s.Fill.ForeColor.SchemeColor = 42
  
    - Listing Scheme Colors
  
      ![ListingSchemeColors](../images/ListingSchemeColors.PNG)
  
      
  
      > sWidth = 20 : sHeight  = 20 : sCols = 8
      >
      > Dim s As Shape
      >     For i = 0 To 80
      >         x = ((i Mod sCols) * sWidth)
      >         y = Int(i / sCols) * sHeight
      >         Set s = Sheet1.Shapes.AddShape(msoShapeRectangle, x, y, sWidth, sHeight)
      >     s.Fill.ForeColor.SchemeColor = i
      > Next i
  
- Removing Fill Colors

  > s.Fill.Visible = msoFalse

- Changing the Brightness (-1 to 1)

  > s.Fill.ForeColor.Brightness = -1

- Tints and Shades

  > s.Fill.ForeColor.TintAndShade = -0.5

- Patterns and Color Gradients

- Changing back color directly is of no effect

  > s.Fill.BackColor.RGB = rgbHotPink

- Changing the Patterns

  ![backColorChanging](../images/backColorChanging.PNG)

  >  With s.Fill
  >
  > ​    .Patterned msoPatternLargeGrid
  > ​    .ForeColor.RGB = rgbLimeGreen
  > ​    .BackColor.RGB = rgbHotPink
  >
  > End With

- Color Gradients

  ![Gradient](../images/Gradient.PNG)

  - OneColourGradient

    > With s.Fill
    >         .ForeColor.RGB = rgbBlueViolet
    >         .OneColorGradient msoGradientDiagonalDown, 1, 1
    >     End With

  - TwoColourGradient

    > With s.Fill
    >         .ForeColor.RGB = rgbBlueViolet
    >         .TwoColorGradient msoGradientHorizontal, 2
    >         .BackColor.RGB = rgbLimeGreen
    >     End With

  - MultipleGradient

    > With s.Fill
    >         .ForeColor.RGB = rgbWhite
    >         .OneColorGradient msoGradientVertical, 1, 1
    >         .GradientStops.Insert rgbRed, 0.25
    >         .GradientStops.Insert rgbGreen, 0.5
    >         .GradientStops.Insert rgbBlue, 0.75
    >     End With



- Formatting Lines

  ![LineFormatting](../images/LineFormatting.PNG)

  > With s.Line
  >         .ForeColor.RGB = rgbRed
  >         .DashStyle = msoLineDash
  >         .Weight = 2.5
  >     End With

- Glow, Reflection, Shadow and Soft Edges

  ![GlowEdgeShadowReflection](../images/GlowEdgeShadowReflection.PNG)

  - glow

    > With s.Glow
    >         .Color.RGB = rgbHotPink
    >         .Transparency = 0.25
    >         .Radius = 15
    >     End With

  - edge

    > With s.SoftEdge
    >         .Type = msoSoftEdgeType1
    >         .Radius = 6
    >     End With

  - shadow

    > With s.Shadow
    >         .Style = msoShadowStyleOuterShadow
    >         .Type = msoShadow25
    >         .Blur = 5
    >         .OffsetX = 8
    >         .OffsetY = 8
    >         .Transparency = 0.25
    >         .ForeColor.RGB = RGB(150, 150, 150)
    >     End With

  - reflection

    > With s.Reflection
    >         .Transparency = 0.25
    >         .Size = 75
    >         .Offset = 3
    >         .Blur = 10
    >     End With

- Basic 3-D Effects

  ![Basic3D](../images/Basic3D.PNG)

  > With s.ThreeD
  >         .Depth = 1
  >         .ContourWidth = 1
  >         .PresetMaterial = msoMaterialMatte2
  >         .PresetLighting = msoLightRigBrightRoom
  >         
  >
  > ​    .IncrementRotationX 25
  > ​    .IncrementRotationY -5
  > ​    .IncrementRotationZ 5
  > ​    
  > ​    .BevelTopType = msoBevelCircle
  > ​    .BevelTopInset = 10
  > ​    .BevelTopDepth = 10
  > ​    .BevelBottomType = msoBevelCircle
  > ​    .BevelBottomInset = 10
  > ​    .BevelBottomDepth = 10
  > End With

- Copying Formats and Setting Defaults  

  - CopyFormatting

    > s1.PickUp
    >     s2.Apply

  - SetingDefaultFormats

    > s1.SetShapesDefaultProperties

  - UsingDefaultFormats

    > s2.ShapeStyle = msoLineStylePreset13
