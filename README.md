# Analyzing Stock Data

## Stock Market

The <u>**purpose**</u> of this project is to help Steve determine which stock he or his parents should invest in. Since Steve has provided us with the stock data of various companies, we will run a VBA script that will automate the process of finding a solution for him. By doing so, Steve will be able to see, with the click of one button, which stock is profitable and which one is not. Also, we will want to evaluate if refactoring our code will have sped up our code execution.

## Analysis and Challenges

First of all, we need to take a look at the data Steve has provided to us. We will be focusing on the <b>Ticker, Daily Volume, and the Yearly Return</b>. This will allow us to determine if the stock is frequently traded and most importantly if the value of the stock has grown over time, making the stock a worthy investment. We will be using VBA to automate the process for Steve, so that he can update his spreadsheets in the future without the need to redo everything. Lastly, we will determine whether or not refactoring code saved us time. 

### <u>Daily Volume and Yearly Return</u>

 After having analyzed the data, we determined that DQ, the stock Steve's parents wanted to invest in, is a bad investment. DQ's total daily volume is rather low and the return is abysmal with a -62.6% return in 2018. 

Now, that we have already written most of the VBA code for analyzing DQ, we want to refactor the code and use it to help us analyze the 12 tickers Steve had previously provided to us in his excel spreadsheet. 
### <u>The refactored code</u>
Firstly, we define our subroutine as "AllStocksAnalysisRefactored." The code execution needs to be timed, in order to do that we need to set a starting and ending point. We will initialize two variables startTime and endTime and set them equal to the VBA's timer function.

```
Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?")

startTime = Timer
```
After doing so, we will create a "tickerIndex" variable that we will set to 0. It is important that the variable is set to 0 since we want to start with the first index in the array being 0 when looping through all the rows in the spreadsheet.

We will not be touching the tickers array nor the header formatted that we created when we first wrote the code.

Three additional arrays will be created as follow:
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

Then let us make sure our tickerVolumes array is initialized to zero. This will make sure that our Daily Volume values are correctly incremented. This needs to be done outside of our next loop, which basically runs the code through all our rows in the spreadsheet.
```
For i = 0 To 11
    tickerVolumes(i) = 0
Next i
```
```
For i = 2 To RowCount

......
......
......

Next i
```

As stated above, we want to increment the Daily Volume. We do this by using our tickerIndex variable as the tickerVolumes array's index. This If statement is located within the "for i = 2 to Rowcount" loop.
```
If Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
End If
```

Now, we make sure that the tickerStartingPrices and tickerEndingPrices arrays also use the tickerIndex variable. The following code basically checks whether the current row ticker and previous row ticker are different. If they are, then the code will return the starting price of the ticker for that given row.
```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then        
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value        
End If
```
We are doing the same with the ending price, but this time we check whether or not the next row matches. On top of that, we tell the VBA to increment the tickerIndex value by one in case the condition below is true.
```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
    End If

End If
```
Finally, we loop through all arrays to output the Ticker, Total Daily Volume, and Return.
```
For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
        
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
Next i
```
Once this is done and we have the correct formatting we need to end the timer and display how long the code took to execute with the following message:
```
endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```

## Results

### <u>Refactoring Code</u>: Advantages
Refactoring code is crucial for keeping code up-to-date. With several people working on your code, the code might become less efficient or messy. By maintaining and refactoring the code you assure that the formatting and readability remains the same. As technology advances, so should your code.
### <u>Refactoring Code</u>: Disadvantages
Refactoring code may be good, however, is it always a good idea to do so?
Not necessarily! For example, if you have to meet a deadline you will not have the time to refactor your code. In this example, you may even introduce unforeseen bugs. Also, for most people the saying goes: "If it ain't broke don't fix it." This applies to many facets of life, but especially code.

Overall, refactoring code, is a great way to assure that the code can be reused or even expanded upon in the future. However, due to the lack of time and budget, most departments usually just tend to use the code that works without worrying about later issues that may arise.

### <u>Our Refactored Code:</u>
The code took significantly less time to run and execute. As can be seen in the two screenshots below. This shows how important refactoring code is. The larger the dataset is, the more time it will take to go through the data, so making sure to write clean and efficient code should always be top priority. 

The execution time of the initial code:

![initial_code](resources/initial_code.png)

vs the refactored code:

![initial_code](resources/refactored_code.png)

It is important to note, however, that while refactoring your code, it is easy to run into execution errors. It may introduce new bugs that could render the code unfunctional. However, since our script is relatively short, we did not encounter such issues.

All in all, make sure to leave notes and comments in your code, since that will help you or anyone else who wants to refactor your code in the future. 


