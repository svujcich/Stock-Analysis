# Stock-Analysis
## Analyzing stock data with the use of VBA
### Overview
Excel is a program that allows the user to store various pieces of information about a given subject, and preform actions, such as calculations, on the data. In many cases, certain actions need to be performed more than once. For example, a calculation might need to be run on multiple data sets, or the header row of a summary table need to be centered in a large, bold font each time a new data set is presented. It is possible to automate these processes using a programming language called VBA, or Visual Basic for Applications. [VBA](https://www.quora.com/Is-VBA-Visual-Basic-for-Applications-considered-a-programming-language) is a dialogue between the user and the program software in which the user writes a list of directions, then VBA runs each statement through an interpreter, and follows the directions step by step . In this analysis, a VBA code that outputs a summary table of stock statistics for a given year is refactored, or improved on, to increase the speed of the subroutine.  Although the original code “allStocksAnalysis” accurately displays the total volume and the return percent for each ticker in a given year, then formats the results with conditional formatting, it is slow to do so. The goal of this analysis is to examine the code logic of a subroutine, and refactor the existing code to loop through the data one time to achieve the same output faster. 

### Results
The first step in the process was to run the original code. Within the code, the Timer function was used to measure the amount of time the computer spent doing the calculations. For both the 2017 and 2018 data sets, it took about 4 seconds to complete the entire subroutine.  

![Screenshot (11)](https://user-images.githubusercontent.com/106559768/176472564-ac1a93be-fc6a-44bf-8b5e-d7e033c8a391.png)
![Screenshot (12)](https://user-images.githubusercontent.com/106559768/176472584-8ff37e49-e317-47a7-9381-3822fd9073d6.png)

  The next step was to examine the code logic. When observing the original code, the subroutine followed sound logic. When the year the user wanted to run the analysis on was entered into an input box, VBA started a timer. It then formatted the output sheet, named the variables and constituent parts of the ticker array, counted the rows, and established that it wanted to perform a loop for every variable in the ticker array. Within that loop, another loop that explained how it wanted the computer to make calculations about the data. At this point, VBA looped through the data for each ticker, line by line, and found the total volume, starting price, and ending price for each ticker. VBA then jumped over to the output sheet and recorded its findings for the given ticker in a summary table. Once the data for the given ticker was input, VBA jumped back to the beginning of the first loop and followed the same directions for each variable in the ticker array. Once all the information for all of the tickers was gathered and recorded, VBA formatted the sheet to make it easier to read, stopped the timer, and displayed a message box (pictured in the previous images) that reported how long the computer spent running the subroutine. 

 Even though this subroutine successfully carried out its task, it contained redundancies. In the original code, VBA jumped back and forth between finding the information for one ticker, and recording the results for that ticker in a summary table. Rather than jumping back and forth between sheets, the refactored code utilized a more efficient way to achieve the same outcome; the refactored code loops through all of the tickers and obtains all the information for all the variables, then, jumps to the output sheet and loops through recording all of the findings for all of the variables. By restructuring the order of tasks and eliminating the need to jump back and forth between sheets, the computer takes fewer steps and uses less memory to run the script.
  
Another improvement was made on the tickerStartingPrices and tickerEndingPrices functions. In the original code, VBA specified two conditions that needed to be met for every line of data for that data to be incorporated in the given ticker’s results; the ticker before (for the starting price) or after (for the ending price) the current ticker in the data set needed to be different from the ticker being analyzed, and the ticker in the row needed to match the ticker that was being analyzed. This is relayed by the code:
         
```
If Cells (j - 1, 1) <> ticker And Cells(j, 1).Value = ticker Then
  startingPrice = Cells(j, 6).Value
  
End If
```

Because the initial loop is already set to run calculations for a given ticker, including a statement that says the ticker needs to match the current ticker is unnecessary. Ultimately, It takes less time to run one condition than two, so removing the statement saying the ticker in the row needed to match the current ticker reduced the amount of time the script runs. 

Finally, a similar situation of redundant instructions presented itself in the conditional formatting section. In the original script, the code specified three conditions for coloring a cell; if the return is positive, the cell turns green, if the return is negative, the cell turns red, and if the return is 0%, the cell remains uncolored. These statements can be reduced to 2 conditions; If the stock return is positive, it turns green, if the stock return is not positive, meaning either 0 or negative, it turns red. This correction can be observed in the following code:

```
For i = dataRowStart To dataRowEnd
        
  If Cells(i, 3) > 0 Then
    Cells(i, 3).Interior.Color = vbGreen
        
   Else
     Cells(i, 3).Interior.Color = vbRed
              
  End If
  
Next i

```

After the adjustments were made to the code, the refactored script was run to test if the changes improved the efficiency of the code. The results were as follows:

![Screenshot (14)](https://user-images.githubusercontent.com/106559768/176479494-acb715f3-5e49-4ea1-bc3b-afbaca608e10.png)

![Screenshot (15)](https://user-images.githubusercontent.com/106559768/176479508-8bfb193f-7975-48ac-aaba-7afe63bc9541.png)

Following the adjustments made to the code, the script ran in about 0.3 seconds. Compared to the original run time of about 4 second, changes made to the scripts increased the speed of the code by about 13 times. 

### Summary
Refactoring a code can present both advantages and challenges. In general, refactoring code is a good practice. For one, the code already exists. Rather than writing a new code from scratch, working with the existing code gives the programmer some framework to improve upon, whereas writing the code from scratch is building the framework. Second, refactoring the code often leads to optimization. Successfully refactored code can increase the speed of a program or generally look more concise on the back end. In the case of this specific project, refactoring the code improved the efficiency of the code while outputting the same results. 

Refactoring code could also be a disadvantage based on the skill level of the programmer. If the programmer does not completely understand the functionality of the software, it is possible to overlook some fundamental principles and encounter errors. From a beginner perspective, when a script runs, it provides instructions for the program. Then, the program follows those instructions to produce a result. If the result is incorrect, the programmer can then go through the script and make corrections that will produce the desired result. As a beginner, I was missing an important, unknown piece of information: [editing a script while the script is running](https://stackoverflow.com/questions/12855750/what-will-happen-when-i-edit-a-script-while-its-running) might confuse the computer.  After a number of unsuccessful attempts to edit the code using correct statements, I made the connection that the debugger was running on a loop. The debugger wanted to make a correction, and move on to the next error until the loop was complete. I also became aware of the pause and stop buttons. Using all of the information gathered plus an [additional resource](https://stackoverflow.com/questions/14903882/freeze-excel-editing-while-vba-is-running) found on Stack Overflow, I started utilizing the stop feature in VBA with success. I was able to stop the program from running and make necessary adjustments, which ultimately led to greater cooperation from VBA.  Without this subtle detail, a beginner might face greater challenges by creating more problems while refactoring a code. 

