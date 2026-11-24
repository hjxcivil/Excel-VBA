### Part 50.2 - Embedding Charts in Worksheets

![copf](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125012150999.PNG)

#### Creating a Chart in a Worksheet

- Using the *AddChart*  *AddChart2* & *Add* Method

  ![ctobj](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125012204736.PNG)


#### Specifying the Chart Type, Position and Size

- Specifying the Chart Type, Position and Size

  ![adad2](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125012220223.PNG)

  

- Selecting all kinds of Data for populating   . . . . .

- Selecting Chart Data Conditionally - - - > *ChartOfRecentFilms*

  ![cojrct](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125012237533.PNG)

#### Referencing the ChartObject and the Chart

- Creating Multiple ChartObjects

  ![Mulitct](../images/Mulitct.PNG)


#### Editing and Formatting the Chart

- Looping Over *ChartObjects*

  ![ctlp](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125012251835.PNG)

- Converting *ChartObject* -- > *Chart* 

  ![coct](../images/coct.PNG)

- Converting Chart Sheets to ChartObjects

#### Controlling the Size and Position

- Positioning and Sizing a ChartObject Based on a Range

  > co.Left = Range("A3").End(xlToRight).Offset(0, 2).Left
  > co.Top = Films.Top

  > co.Height = Films.Height
  > co.Width = Films.Width * 3

- Move & Position Multiple Charts to One Sheet

  ![Mcctol](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125012305348.PNG)

  