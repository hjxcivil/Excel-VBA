### Cart 28.1 - Creating PowerPoint Presentations

[TOC]

#### Creating Presentations and Slides

- Creating Presentations

  ![presadd](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014757820.PNG)

  > Dim ppPres As PowerPoint.Presentation
  >     Set ppPres = ppApp.Presentations.Add
  
- Creating a Title Slide

  ![addSld](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014810832.PNG)
  
- Adding Text to Textboxes

  ![sldatt](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014826040.PNG)
  
  *.TextRange.Text*

#### Copying Tables into PowerPoint

- Copying an Excel Range into PowerPoint

  ![sldexc](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014840800.PNG)

  > ppSlide.Shapes.PasteSpecial ppPasteOLEObject

- Adding a Custom Textbox and Formatting

  ![ppcstbox](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014852750.PNG)

  


#### Copying an Excel Chart into Power Point

![sldct](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014904677.PNG)



#### Applying a Template  or Theme to a Presentation

![tplthm](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014916923.PNG)

#### Saving and Closing a Presentation & Late binding

![savecloselate](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125014929734.PNG)
