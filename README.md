# Stock-Analysis

## Stock Analysis With Excel VBA

## Overview of Project
The purpose of this project was to analyze what refactoring existing Microsoft Excel VBA code to determine if it would speed up the processing time.  The existing code analyzed a sample of stocks for the years 2017 and 2018 to determine for the client if the stocks were worth investing in.  Specifically analyzing the total daily volume of each stock and what the overall rate of return was in a (+/-)  (%) format.

### Analysis Results 
Project began with utilizing the Excel VBA code that was provided to create each input box, chart headers, as well as the stock ticker array, and lastly to activate the required worksheet “AllStocksAnalysis”.  Results from 2017 show that the all stocks analysis yielded a majority of positive rates of return, with only one stock TERP performing at a loss of 7.2% of TERP stock.  The highest performing stock in 2017 was DQ at a rate of return of 199.4%.  Results from 2018 were drastically different then 2017 with the majority of stocks yielding negative rates of return.  Only two stocks from 2018 yielded positve returns of 81.9% (ENPH) and 84.0% (RUN).

[Refactored code.docx](https://github.com/dianahcortez/Stock-Analysis/files/8977638/Refactored.code.docx)

<img width="500" alt="2017 All Stocks" src="https://user-images.githubusercontent.com/104927745/175658791-76a72364-beff-42b5-9ba1-24f179a53987.png">

<img width="500" alt="2018 All Stocks" src="https://user-images.githubusercontent.com/104927745/175658808-aea3fcb9-9508-4cc4-938a-3092b1124067.png">

### Summary
As stated by bmc.com, refactoring code is the process of restricting computer code without changing or adding to its external behavior and functionality.  Stating more simply, its turning dirty or “bulky” code into a cleaner and more efficient format.
[Source] https://www.bmc.com/blogs/code-refactoring-explained/

### The Pros and Cons of Refactoring Code
As stated previously, refactoring code helps us clean it up and make it more organized.  Advantages of doing this refactoring process is to speed up the software, assist with debugging and makes for a more organized and faster programming.  Another advantage is a refactored code is much easier to read, because it is concise and forthright.  I noticed the biggest advantage to refactoring this code was that it dramatically decreased the macro run time.  The run time of the non refactored macro took 0.476 seconds (2017) and 0.445 seconds (2018).  The run time of the refactored code was decreased to 0.070 (2017) and 0.056 (2018).

<img width="500" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104927745/175659543-493cbf78-9eba-4636-85c3-37a0070f1ebd.png">

<img width="500" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104927745/175659556-32660f2d-84cb-4853-a15b-473261d83fd5.png">

Some disadvantages of refactoring could be when one has a very large application, which could lead to new bugs and/or errors in the code.

### Challenges and Difficulties Encountered
Challenges I faced with this project was identifying and defining the arrays as it was a very new concept to me.  VBA is extremely finicky with even the smallest mistake breaking a whole ‘If Then code.  I struggled with making sure the code was properly formatted/clean, and was able to find all my typo’s by looking at the code line by line to see where the breaks were.  Another challenge I ran into was sometimes the timer would display scientific notation instead of seconds.  Seems to be at random.

<img width="500" alt="Challenge Image" src="https://user-images.githubusercontent.com/104927745/175658158-bfe0ac16-8fef-4bc2-928d-e4de4f24ea61.png">

