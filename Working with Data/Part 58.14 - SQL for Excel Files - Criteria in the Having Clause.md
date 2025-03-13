## Part 58.14 - SQL for Excel Files - Criteria in the Having Clause

![Having](../images/Having.PNG)

#### Creating Groups and Aggregates

![havrecap](../images/havrecap.PNG)

#### Aggregation Functions in the Where Clause

> WHERE [Oscar Wins] > 0 AND Count([Title]) >= 10  - > SYNTAX ERROR !!!!

![wheragg](../images/wheragg.PNG)

#### Criteria in the Having Clause

> HAVING Count([Title]) >= 10 

![hav](../images/hav.PNG)

#### Multiple Criteria in the Having Clause

![havmu](../images/havmu.PNG)

#### Controlling Multiple Criteria

> "HAVING  (Count([f.Title]) >= 10 OR Sum(f.[Oscar Wins] >= 20)) AND Avg(f.[Run Time]) >= 150 
