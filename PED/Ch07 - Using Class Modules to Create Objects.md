# Ch07 - Using Class Modules to Create Objects

[TOC]

##  Create Class Objects

- Analyze single cell:(Empty ,Label, Constant, formula)

  ![cc1](image/cc1.PNG)

- Create the CCell:

  ![cc2](image/cc2.PNG)

  -  public variables 

    > anlCellType muCellType mrngCell

  - Property Procedures

    > Cell CellType DescriptiveCellType

  - Methods

    > Analyze

  - The Analyze Method of the Cell Object

    > Analyze = Me.DescriptiveCellType
    >
    > MsgBox clsCell.Analyze()

##  Creating a Collection

- Using Collection to  analyze a worksheet or ranges of cells

  ![ccol](image/ccol.PNG)

-  Access a specific Cell object

  > Set clsCell = gcolCells(3) Set clsCell = gcolCells(“$A$3”)

## Creating a CCells Object

- highlight cells of the same type and another method

  ![cc3](image/cc3.PNG)

- Update the Ccell

  ![cc4](image/cc4.PNG)

- Create the Ccells

  ![cc5](image/cc5.PNG)

  - The two ShortComings

    - *Can not Use For...Each to process the members*
    - *has no default property: gclsCells(1) is not permitted*

  - Using a Text Editor to Solve

    ![cc6](image/cc6.PNG)

    - export Ccells - > edit -> import

      ![mke](image/mke.PNG)

##  Trapping Events

- respond to events

  ![tra1](image/tra1.PNG)

  - Declare a WithEvents variable in a class module.

    > Private WithEvents mwksWorkSheet As Excel.Worksheet

  - Assign an object reference to the variable.

  - > Property Set Worksheet(

  - Additions to the *CCells* Class Module

    ![cls](image/cls.PNG)

##  Raising Events

- define own events and trigger them in the code

  - The Cells raises an event that will be trapped by the Cell objects
    - An Event declaration at the top of the class module
    - A line that uses RaiseEvent to cause the event to take place

  - Changes to the CCells Class Module to Raise an Event

    ![ccls2](image/ccls2.PNG)

    - Move the Enum anlCellType from CCell

    - Event declaration

      ```
      Event ChangeColor(uCellType As anlCellType, bColorOn As Boolean)
      ```

    - RaiseEvent

      ```
      RaiseEvent ChangeColor(mcolCells(Target.Address).CellType, True)
      ```

    -  created an explicit parent-child relationship

      > Set clsCell.Parent = Me

  - Changes to the CCell Class Module to Trap the ChangeColor Event

    ![ccs3](image/ccs3.PNG)

- A Family Relationship Problem

  -  multiple times creates a memory leak

    - Create the Terminate Method in CCell

      > Public Sub Terminate()
      >     Set mclsParent = Nothing
      > End Sub

  - two objects that store references to each other

    - Create the Terminate Method in CCells

      > Dim clsCell As CCell
      >     For Each clsCell In mcolCells
      >         clsCell.Terminate
      >     Next clsCell
      >     Set mcolCells = Nothing

  - Update *CreateCellsCollection* Procedure

    ![image-20230212211137645](../../AppData/Roaming/Typora/typora-user-images/image-20230212211137645.png)

## Creating a Trigger Class

- Creating a trigger class to raising the ChangeColor event 

  ![ct](image/ct.PNG)

- Update the CCells & CCell (Remove Terminate ...)

- Trap the ChangeColor Event of CTypeTrigger

  ![cc9](image/cc9.PNG)

- Changes  the CCells Assign References to CTypeTrigger to Cell Objects

  ![ccs31](image/ccs31.PNG)

  