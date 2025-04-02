MY PROJECTS
===========

In this, I listed all the projects I build in Excel, Vba, Power Query, Python and many more .... only Videos

## 1) Excel Dynamic Header
In this project, I use some formulas to change the header dynamically on selecting ENTITY_NAME. As well as, If any column having listed item then on that particular column, the dynamic dropdown list will be shown.

**Header**

[![WATCH Excel Dynamic Header](https://github.com/deepumondal1/MyProjects/blob/master/videos/UPLOAD%20Automation_COMPRESS.gif)](https://github.com/deepumondal1/MyProjects/blob/master/videos/UPLOAD%20Automation_COMPRESS.mp4)

**Dynamic Dropdown**
[![WATCH Excel Dynamic Header](https://github.com/deepumondal1/MyProjects/blob/master/videos/UPLOAD%20Automation_COMPRESS_2.gif)](https://github.com/deepumondal1/MyProjects/blob/master/videos/UPLOAD%20Automation_COMPRESS.mp4)


## 2) Excel Item Splitter
In this project, I use VBA to split the item and keep entry on different fields on selecting the items.

[![WATCH Excel Dynamic Header](https://github.com/deepumondal1/MyProjects/blob/master/videos/Item%20Project2.gif)](https://github.com/deepumondal1/MyProjects/blob/master/videos/Item%20Project2.mp4)


## 3) Excel Row Transformation
In this project, I use power query to combine each row in single column to make actual table header and split it to get the visible table.

#### Example :-

**Raw Data :-**

| Column A |
|:-|
| **A#ItemCode#Name#Date#** |
| **B#Pos#Seq#** |
| **C#Qty#Amnt#Unit#** |
| **A**#001#Apple#01-01-2025# |
| **B**#5#1# |
| **C**#10#1200#kgs# |
| **A**#002#Banana#01-01-2025# |
| **B**#10#1# |
| **C**#24#160#dzn# |
| **A**#002#Orange#01-01-2025# |
| **B**#15#1# |
| **C**#100#2000#kgs# |
| **A**#002#Coconut#01-01-2025# |
| **B**#20#1# |
| **C**#50#150#pcs# |


**Convert to :-**

| ItemCode | Name | Date | Pos | Seq | Qty | Amnt | Unit |
|:-|:-|:-|:-|:-|:-|:-|:-|
| 001 | Apple | 01-01-2025 | 5 | 1 | 10 | 1200 | kgs |
| 002 | Banana | 01-01-2025 | 10 | 1 | 24 | 160 | dzn |
| 003 | Orange | 01-01-2025 | 15 | 1 | 100 | 2000 | kgs |
| 004 | Coconut | 01-01-2025 | 20 | 1 | 50 | 150 | pcs |

[![WATCH Excel Dynamic Header](https://github.com/deepumondal1/MyProjects/blob/master/videos/ExcelTransform.gif)](https://github.com/deepumondal1/MyProjects/blob/master/videos/ExcelTransform.mp4)



## 4) Excel 2D to D Transformation
In this project, I use a single line array function in excel to restructure the table into single column.

#### Example :-

**Raw Data :-**
| ITEM | A | B | C | 
| :-: | :-: | :-: | :-: |
| APPLE | 100 | 200 | 300 | 
| ORANGE | 400 | 500 | 600 | 
| BANANA | 700 | 800 | 900 | 
| Grape | 1000 | 1100 | 1200 | 
| Coconut | 1300 | 1400 | 1500 | 

**Convert to :-**
| Column |
| :- |
| APPLE | 
| 100 | 
| 200 | 
| 300 | 
| ORANGE | 
| 400 | 
| 500 | 
| 600 | 
| BANANA | 
| 700 | 
| 800 | 
| 900 | 
| Grape | 
| 1000 | 
| 1100 | 
| 1200 | 
| Coconut | 
| 1300 | 
| 1400 | 
| 1500 | 


[![WATCH Excel Dynamic Header](https://github.com/deepumondal1/MyProjects/blob/master/videos/Excel%202D_to_D.gif)](https://github.com/deepumondal1/MyProjects/blob/master/videos/Excel%202D_to_D.mp4)
