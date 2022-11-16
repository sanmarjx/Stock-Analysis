# Stock-Analysis With Excel VBA
## Overview of Project
### Purpose
The purpose of this project is to refactor a Microsoft Excel VBA code to obtain stock information from the years 2017 and 2018 and determine whether or not the stocks were a good investment. We want to refactor the code to make it more efficient by making it easier to read and taking fewer steps. We will be able to test if our code is running faster by checking the run time for each Macro.
### Results
#### Analysis
The project began with a code that we copied and pasted into VBA. We then refactored the code to make it more efficient. This is the process that we followed to improve the code.
##### 1- Created a ticker index and set it to 0. We also created the output arrays for the volume, starting price and ending price.
![Step1 1](https://user-images.githubusercontent.com/116690861/201403978-50fdf6df-5b42-4cda-8b0c-3cda3cee10b8.png)

##### 2- Created a loop to initialize the volume to zero. Created a second loop to go over all the rows. 
- In this step I also set the starting and ending price to zero because I thought it would be beneficial to the code.
![Step 2](https://user-images.githubusercontent.com/116690861/201404461-c119bca4-540c-4e31-b4f9-c6723b61465a.png)

##### 3- Included IF statements to the loop find the start and end prices of the tickers as well as increasing the ticker index.
![Step 3](https://user-images.githubusercontent.com/116690861/201406012-01b70485-04a3-493b-a4e9-b2be87dcfca2.png)

##### 4- Looped through the arrays to output the ticker, volume and return.
![Step 4](https://user-images.githubusercontent.com/116690861/201408167-ced80465-5756-4564-8660-7b24e6fa274a.png)

##### 5- Compared script execution times
The execution times of the script were faster than before the refactoring which means that we made the script more efficient.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/116690861/201409057-6dfe1e06-7e20-4ac8-9740-2f8489f3ebbf.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/116690861/201409107-e0e56f67-f80d-4067-afe8-bc9105d07cf3.png)

### Summary
#### Advantages and Disadvantages of Refactoring Code
Refactoring helps us organize our code better. We can find better ways to write the code to perform the same tasks. It can make the script easier to read by simplifying some of the code and helping us find mistakes and redundancies quicker. However, there are also some disadvantages. It can take a long time to try to refactor the code and then you might not find any way to improve it. Also, refactoring can lead to new bugs and errors that were not in the code before. 

#### Advantages and Disadvantages of Refactoring Our Stock Analysis
The biggest advantage that we saw in our dataset was the improved run time. In addition, a cleaner code will help us in the future if we decide to add to this project. If it is easier to read, we will be able to pick up right where we left much faster than with a messy code.
