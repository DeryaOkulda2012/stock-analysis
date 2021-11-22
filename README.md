Module 2 - Challenge

VBA Code Refactoring and Stock Analysis

#Overview:

Steve can analyze an entire dataset at the click of a button. To do more research for his parents, Steve wants to expand the dataset to include the entire stock market over the last few years. Although the initial code he has works well for the dozen stocks, it will need to be expanded for thousands of stocks. And if it does, it may take a long time to execute.

In this module report the code is edited, or refactored, to loop through all the data one time in order to collect the same information in this module. Refactoring will be used to make the code more efficient — by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Although much of the functionality could be done directly in MS Excel to provide the same data analysis, using VBA Script for the same functionality can significantly improve the overall efficiency of the computer and memory.

#Results:

Large set of market data for a dozen different stocks collected in 2017 and 2018 was analyzed to compare the two-year performances of the twelve stocks.

##Market Analysis 2017

The twelve stocks analyzed in this dataset clearly performed much better in 2017 than they did in 2018. In fact, in 2017 eleven of the twelve stocks performed really well, including two which more than doubled (ENPH with an increase of 129.5% and FSLR of 101.3%) and two stocks which nearly tripled in market value (DQ at 199.4% and SEDG at 184.5%). Only one stock lost value with a slight loss of 7%. 
See figure ‘Returns 2017.png’

The code took 0.09204 seconds to analyze the 2017 market data for the twelve stocks.
See figure ‘VBA_Challenge_2018.png‘

##Market Analysis 2018

The same stocks performed much more poorly in 2018 than they did in 2017, some losing nearly half their market value (DQ losing 62.6%, JKS losing 60.5%, and SPWR 44.6%). Only two stocks performed well, and in fact quite well, nearly doubling their market value – ENPH gaining 81.9%, and RUN gaining 84%.

Of all stocks, ENPH performed spectacularly gaining 129.5% in 2017 and 81.9%, for a cumulative 417% market gain! Too bad we don’t have a glass ball to see the future.
See figure ‘Returns 2018.png’

This time, the code took just a bit longer, at 0.09411 seconds to analyze the 2018 market data for the twelve stocks.
See figure ‘VBA_Challenge_2017.png‘

#Summary:
There is a definite advantage in refactoring (or editing) code in general, and VBA Script in particular, especially when analyzing large chunks of data. In this particular example, several thousand rows of daily market data was analyzed.

The most obvious advantage is that refactoring leads to better quality code. Refactoring allows for simplified support and code updates, improved code readability and reduced complexity. Reduced complexity for easier understanding, as well as maintainability and scalability. Cleaner code is much easier to update and improve. Ultimately, it helps to save time and money in the future.

The advantage of refactoring VBA Script is it allows for better analysis of the data. Although much can be done directly in MS Excel which could provide the same analysis, using VBA script code for the same functionality improves the overall efficiency of the computer significantly.

The potential disadvantage is that you might have to retest lots of functionality. This was obvious as more of the functionality was incorporated into the VBA script and additional testing and re-testing needed to be performed to make sure the code functioned properly.

The more code you write, the more possibility of running into bugs which will need to be cleaned up through multiple development iterations and testing. The larger the size of the data and bigger the operation, the riskier the refactoring will be. It all comes down to how much money is available for refactoring and how much time.

