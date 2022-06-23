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

Steve's primary request to us was to help him to automate the analysis process as much as possible. While it's fairly straightforward to manually run an analysis for a single stock, it makes sense to develop a system that speeds up the process if you are going to be repeatedly running the same analysis on a large number of stocks. Therefore, we developed a set of Visual Basic for Applications (VBA) code that would quickly analyze any number of stocks extremely quickly. The code was initially created for Steve's immediate requirement to analyzed the dozen green energy stocks he was focused currently on and it worked very well for that purpose. We soon realized however, that with a little refactoring, the code could run much faster and he could apply it to any amount of stock data. Even a small improvement in efficiency would save Steve a lot of time in the long run.

# Results

## Stock Performance

For this project, we ended up analyzing the performance data for a dozen stocks over a two year period. We looked at 2017 and 2018 separately to look for any obvious trends from one year to the next. (Expanding this dataset to several years before and after 2017-2018 would be a next logical step). For each stock, we summarized two key pieces of information:

- **Total Daily Volume** to quantify the amount of trading that took place involving each stock
- **Stock Return** to see the percent rate of return over the course of year

Even looking casually at the summary tables below, one can conclude that investing in any of these stocks could be a risky proposition as best. After a stong collective 2017, all but two stocks displayed a negative return in 2018. The most promising of these is "RUN", which saw a return of 84.0% in 2018 after posting a modest return of 5.5% in 2017. Additionally, total volume increased by almost 88% from one year to the next. The other stock with a positive return in 2018 was "ENPH", displaying a return of 81.9% in 2018 after posting a significant return of 129.5% in 2017. However, the total volue of shares traded was a massive 607,473,500, amost three times more than in 2017. 

**2017**

![Image of 2017 Stock Analysis table](https://github.com/brianbutler08/stock-analysis/blob/main/Stock%20Analysis%202017.png?raw=true)

**2018**

![Image of 2018 Stock Analysis table](https://github.com/brianbutler08/stock-analysis/blob/main/Stock%20Analysis%202018.png?raw=true)

## Code Performance

Our initial set of code did a great job of pulling out and summarizing the information discussed above. It looped through the data set multiple times to find the appropriate data and push it to our summary table. However, in order to improve its performance on larger datasets and allow Steve to analyze potentially thousands of stocks quickly, we needed to refactor the code. Th primary approach we took was to set up a series of four arrays, one for the tickers and three for the outputs - volume, starting price and ending price. This meant that we would not need to loop through the data nearly as many times, resulting in a more streamlined code and much faster processing times.

![code image](https://github.com/brianbutler08/stock-analysis/blob/main/Screen%20Shot%202022-06-22%20at%2011.06.13%20PM.png)



# Summary
