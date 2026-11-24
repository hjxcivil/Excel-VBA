### Part 58.14 - SQL for Excel Files - Criteria in the Having Clause

![Having](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125024845870.PNG)

#### Creating Groups and Aggregates

![havrecap](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125024855687.PNG)

#### Aggregation Functions in the Where Clause

> WHERE [Oscar Wins] > 0 AND Count([Title]) >= 10  - > SYNTAX ERROR !!!!

![wheragg](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125024906474.PNG)

#### Criteria in the Having Clause

> HAVING Count([Title]) >= 10 

![hav](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125024916920.PNG)

#### Multiple Criteria in the Having Clause

![havmu](https://cdn.jsdelivr.net/gh/hjxcivil/image-hosting@main/img/20261125024956146.PNG)

#### Controlling Multiple Criteria

> "HAVING  (Count([f.Title]) >= 10 OR Sum(f.[Oscar Wins] >= 20)) AND Avg(f.[Run Time]) >= 150 

