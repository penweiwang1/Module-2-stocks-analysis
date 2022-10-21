# Deliverable 3

## Overview of Project: 
The purpose of this analysis is to refactor VBA code via excel from the provided stock data  to see if it could improve the efficency of the original code and ultimately determine which stock is worth investing.

## Results: Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.
<img width="244" alt="Screen Shot 2022-10-20 at 9 32 08 PM" src="https://user-images.githubusercontent.com/115126898/197090510-b5bb815d-227a-490f-a858-8346d8f710e3.png">
<img width="244" alt="Screen Shot 2022-10-20 at 9 32 38 PM" src="https://user-images.githubusercontent.com/115126898/197090530-5a858afa-29a7-4f61-a4fd-6356855ed0dc.png">

As seen in the images above, 2017 was a good year for stocks as most of the stocks had a positive return with the exception of "TERP". However, in 2018 all but two stocks ("ENPH" and "RUN") had a negative return. Therefore stocks "ENPH" and "RUN" are worth investing as both had consistent positive returns for two consecutive years. 

The run time of the original code for the year 2017 compared to the refactored code decreased more than half, from 0.289 seconds to 0.070 seconds. Similarly, run times decreased from that of the original code to the refactored code for the 2018 data. Below are the run times for the refactored code for years 2017 and 2018, both of which are very quick.

<img width="270" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/115126898/197095537-3956b4de-7984-496a-a500-8b5a285f288f.png">
<img width="270" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/115126898/197095544-ee59aecd-652e-4d8d-b080-d98e1bf0e7c1.png">

Decrease in run times is may be attributed to the shortened number of loops the system has to run. In the image below from the original code, there are two combined loops which goes through (For i = 0 to 11) and (For j = 2 to rowcount), meaning the code runs about 12 x 12 times.
<img width="487" alt="Screen Shot 2022-10-20 at 10 00 20 PM" src="https://user-images.githubusercontent.com/115126898/197093507-0f3a36b8-5f71-4482-b0e6-cbb6cf49bf5d.png">

Meanwhile for the refactored code, see image below, the same corresponding code is written so the system only has to go through one loop (For j = 0 to Rowcount). 
<img width="652" alt="Screen Shot 2022-10-20 at 10 03 21 PM" src="https://user-images.githubusercontent.com/115126898/197093929-40b93085-4a56-4e6a-95c4-35b35a239bef.png">

## Summary:
### What are the advantages or disadvantages of refactoring code?
Some advantages of refactoring code is that it may help reduce the run time of the program. This is because the code has been revised to be more concise and easier to read, which may also help with debugging. Some disadvantages is that sometimes it could be time consuming, especially if the data file is very large. New bugs can also be introduced when rewriting the code. Another disadvantage is when a different user refactors code and  interprets the original code differently,  the end result might not serve the same purpose that the original user intended. 

### How do these pros and cons apply to refactoring the original VBA script?
An advantage of the the refactored VBA script was that it helped decrease the run time by more than half compared to the original script. From 0.289 seconds to 0.070 seconds for year 2017. A disadvantage was that it introduced a new variable (tickerIndex) that for a inexperienced coder, it can increase the chance of creating a bug if not used correctly. 



Results
The analysis is well described with screenshots and code (4 pt).

