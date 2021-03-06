# stock-analysis

# Overview of Project

This project began when our friend Steve first contacted us to assist him in analyzing some stocks. He had recently graduated with his finance degree and had taken on his parents as his first clients. They were interested in investing in DAQO New Energy Corp, a company working in the renewable energy market, but knew very little about them. Steve approached us to give him a hand in analyzing the stock performance history of DAQO, as well as a number of other green energy entities. 

Steve has a little experience with Microsoft Excel and has compiled stock data from twelve different companies, from two different years (2017 and 2018). For each stock, the dataset contains information for each trading day of the year, including:

- Date of trading
- Stock value at opening
- High value for the day
- Low value for the day
- Stock value at close
- Adjusted value at close
- Stock trading volume

Steve's primary request to us was to help him to automate the analysis process as much as possible. While it's fairly straightforward to manually run an analysis for a single stock, it makes sense to develop a system that speeds up the process if you are going to be repeatedly running the same analysis on a large number of stocks. Therefore, we developed a set of Visual Basic for Applications (VBA) code that would quickly analyze any number of stocks extremely quickly. The code was initially created for Steve's immediate requirement to analyze the dozen green energy stocks he was focused currently on and it worked very well for that purpose. We soon realized however, that with a little refactoring, the code could run much faster and he could apply it to any amount of stock data. Even a small improvement in efficiency would save Steve a lot of time in the long run.

# Results

## Stock Performance

For this project, we ended up analyzing the performance data for a dozen stocks over a two year period. We looked at 2017 and 2018 separately to look for any obvious trends from one year to the next. (Expanding this dataset to several years before and after 2017-2018 would be a next logical step). For each stock, we summarized two key pieces of information:

- **Total Daily Volume** to quantify the amount of trading that took place involving each stock
- **Stock Return** to see the percent rate of return over the course of year

Even looking casually at the summary tables below, one can conclude that investing in any of these stocks could be a risky proposition as best. After a strong collective 2017, all but two stocks displayed a negative return in 2018. The most promising of these is "RUN", which saw a return of 84.0% in 2018 after posting a modest return of 5.5% in 2017. Additionally, total volume increased by almost 88% from one year to the next. The other stock with a positive return in 2018 was "ENPH", displaying a return of 81.9% in 2018 after posting a significant return of 129.5% in 2017. However, the total volume of shares traded was a massive 607,473,500, almost three times more than in 2017. 

**2017**

![Image of 2017 Stock Analysis table](https://github.com/brianbutler08/stock-analysis/blob/main/Stock%20Analysis%202017.png?raw=true)

**2018**

![Image of 2018 Stock Analysis table](https://github.com/brianbutler08/stock-analysis/blob/main/Stock%20Analysis%202018.png?raw=true)

## Code Performance

Our initial set of code did a great job of pulling out and summarizing the information discussed above. It looped through the data set multiple times to find the appropriate data and push it to our summary table. However, in order to improve its performance on larger datasets and allow Steve to analyze potentially thousands of stocks quickly, we needed to refactor the code. Th primary approach we took was to set up a series of four arrays, one for the tickers and three for the outputs - volume, starting price and ending price. This meant that we would not need to loop through the data nearly as many times, resulting in a more streamlined code and much faster processing times.

![code image](https://github.com/brianbutler08/stock-analysis/blob/main/Screen%20Shot%202022-06-22%20at%2011.06.13%20PM.png)

![code image 2](https://github.com/brianbutler08/stock-analysis/blob/main/Screen%20Shot%202022-06-22%20at%2011.07.22%20PM.png)

![code image 3](https://github.com/brianbutler08/stock-analysis/blob/main/Screen%20Shot%202022-06-22%20at%2011.07.45%20PM.png)

In order to evaluate the time required to run each set of code (original and refactored) we added a timer which began as the code was run and ended as the code finished. The original code ran quite quickly in looping through the twelve tickers in between 0.32 and 0.52 seconds. However, once the code was refactored, it ran much faster. The code handled the 2017 data in around 0.234 seconds and the 2018 data in a remarkable 0.086 seconds. 

![2017 timer](https://github.com/brianbutler08/stock-analysis/blob/main/VBA_Challenge_2017.png?raw=true)

![2018 timer](https://github.com/brianbutler08/stock-analysis/blob/main/VBA_Challenge_2018.png?raw=true)

# Summary

## Why Do We Refactor Code?

In general, code is refactored to make inefficent, messy and over-complicated code cleaner and more efficicent. By clearing out unecessary lines of code and restructuring how it is organzied, the application becomes faster, easier to maintain and simpler to understand. Coding is often a collaborative process, with multiple people working together on a piece of code over a period of time. For this process to work, it is critical to maintain clearly understandable and straightforward code. Keeping code efficient and streamlined also makes it more scalable. The primary disadvantage to refactoring is always the potential of breaking the original code. In an effort to make it more efficient, you may accidently remove or modify a critical element and stop it from running properly.

## Refactoring VBA Code on This Project

Our friend Steve was more than happy with our original code and didn't really care or understand about how it worked or how efficient it was. However, once there was an interest in using it to analyze much larger datasets, we realized that it would benefit greatly from a refactoring. The major advantage would or course be an increase in speed when moving through large data sets. Additionally, the code would be easier to scale up or modify if Steve wanted to look at additional qualities of differents stocks. As stated above, the major disadvantage here would be modifying the code to the extent that it breaks. This scenario indeed happened to us in refactoring the code for this project and it took some work to get the new refactored code to work again. As a general practice, we kept a copy of the original code to work from so always had a set of working code for Steve to run.
