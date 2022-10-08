# stock-analysis
## Using VBA to analyze stock market data.

***
### Overview
#### I've leverged the functionality of VBA to analyse stock market data from two excel tables, each table contains thousand of rows and are seperated by year(2017 & 2018). The script I created takes user input to select the year of analysis, loops through all rows of a given year (sorted by the stock's unique "ticker" name), and analyizes the yearly return of each stock. 
- Yearly return was calculated as the change from opening price at the start of a given year to the closing price at the end of that year 
- The cells are conditionally formatted so that a cells with a positive yearly change are green and cells a negative yearly change are red
- The percent change from opening price to closing price is also displayed

#### In addition to this analysis, I have refactored my code to make it more efficient in processing the large excel dataset. 

***
### Results: Stock Performance
#### 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/105818879/194681528-b4b58b34-2077-47d9-acb0-9af3744b0fd6.png)
#### 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/105818879/194681535-cb798a79-6d22-49d2-9c0b-7cbe64733b35.png)

#### Using conditional formatting and displaying the "return" in a percent format allows us to easily compare performance across individual stocks as well as across different years. The data shows that 2017 was generally a good year for the stocks that have been analyzes and 2018, in contrast, was generally a bad year for their performance. ENPH and RUN were the only two stockp that performed positively across both years and TERP was the only stock to perform negatively across both years.

***
### Results: Refactored Script vs. Original Script 
#### The refactored script seen in the section above above vastly outperformed the original script. Run time, compared to the original, has been greatly reduced- From ~.77 seconds down to .14 seconds.
- See original runtimes below:
#### 2017
<img width="539" alt="2017_unRF" src="https://user-images.githubusercontent.com/105818879/194683755-3474516d-1bc1-46ec-bab9-cc7a2385c511.png">
#### 2018
<img width="556" alt="2018_UnRF" src="https://user-images.githubusercontent.com/105818879/194683762-94f4cee3-ca36-4c14-bd30-68bdfaaa635a.png">


***
### Summary:
#### Refactoring offers several significant advantages. It can improve efficiency by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. The disadvantge is that imprecise refactoring has the potential to introduce new bugs and errors into the code.
