## How do I paste and resize a picture in Word with Excel VBA

![wdps](../images/wdps.PNG)

#### Creating a Word Document with Text

![wdbsc](../images/wdbsc.PNG)

#### Pasting Excel Data as a Word Table

> wd.Selection.Paste 

#### Pasting Excel Data as a Picture

- The PasteSpecial Method

  > .PasteSpecial 'table default

  - Pasting Excel Data as a Picture 

    > .PasteSpecial DataType:=wdPasteEnhancedMetafile 

    - Setting the Placement of a Picture

      > .PasteSpecial ..., Placement:=wdInLine 'Default cann't drag to move
      >
      > .PasteSpecial ..., Placement:=wdFloatOverText 'can drag to move

  - Returning the Current Width and Height

    > Debug.Print doc.Shapes(1).Width, doc.Shapes(1).Height

  - Resizing a Floating Shape

    > doc.Shapes(1).Width = 400

  - Resizing an Inline Shape

    > doc.InlineShapes(1).Width = 400
