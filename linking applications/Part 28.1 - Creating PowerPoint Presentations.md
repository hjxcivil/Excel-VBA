### Cart 28.1 - Creating PowerPoint Presentations

[TOC]

#### Creating Presentations and Slides

- Creating Presentations

  ![presadd](../images/presadd.PNG)

  > Dim ppPres As PowerPoint.Presentation
  >     Set ppPres = ppApp.Presentations.Add
  
- Creating a Title Slide

  ![addSld](../images/addSld.PNG)
  
- Adding Text to Textboxes

  ![sldatt](../images/sldatt.PNG)
  
  *.TextRange.Text*

#### Copying Tables into PowerPoint

- Copying an Excel Range into PowerPoint

  ![sldexc](../images/sldexc.PNG)

  > ppSlide.Shapes.PasteSpecial ppPasteOLEObject

- Adding a Custom Textbox and Formatting

  ![ppcstbox](../images/ppcstbox.PNG)

  


#### Copying an Excel Chart into Power Point

![sldct](../images/sldct.PNG)



#### Applying a Template  or Theme to a Presentation

![tplthm](../images/tplthm.PNG)

#### Saving and Closing a Presentation & Late binding

![savecloselate](../images/savecloselate.PNG)
