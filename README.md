# Stock-Analysis

## Overview of Project
We've created a workbook where a dataset of stocks can be analyzed with a click of a button. The purpose of this project is to reactor this code by updating the VBA script so the analysis can run faster. 

## Coding Steps
-	The initial file has a VBA module “All Stock Analysis” where the output was displayed in the All Stock analysis tab. We’ve created a new module name “VBA_Challenge” and the analysis will be displayed in the VBA_Module spreadsheet we created.

-	In the previous module we analyze the data by sorting through the data by variable and then iterate through the entire dataset to find the totalVolume for each ticker:
‘Dim startingPrice As Single
‘Dim endingPrice As Single

-	In the refactored code we created separate output put arrays and stored the date for each ticker result:
' Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

## Result
-	After running the both module the result shows the new refactored code in VBA_Challenge module takes less time to display the same output. Below are the results screenshot:

![All_Stock_Analysis Module Runtime 2017](https://user-images.githubusercontent.com/85530486/127730118-7ebe2ff2-0887-4fb5-b2d8-45d668199a89.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/85530486/127730124-81ca5e2b-84a1-4084-a7f8-05bc26b8e63c.png)
![All_Stock_Analysis Module Runtime 2018](https://user-images.githubusercontent.com/85530486/127730128-ca3759fc-8eba-4c19-baa2-d54cc0c634c4.png)
![VBA_Challenge_ 2018](https://user-images.githubusercontent.com/85530486/127730130-841dad4f-528e-4a63-97a8-e903b6ed7446.png)

## Summery
-	Refactoring is a key part of the coding process. By refactoring this code, didn’t add new functionality but made the code more efficient—by taking fewer steps, using less memory and improving the logic of the code to make it easier for future users to read.
-	As we can see on this project the refactored code on the VBA_Challenge module made the script more efficient by taking fewer steps, using less memory, thus providing the output faster
