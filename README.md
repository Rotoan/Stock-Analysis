### VBA stock analysis challenge
( ˇ෴ˇ )

#Purpose
	This project was intended to use VBA and excel to analyze two years of stock data and use this to assist
a pair of amateur investors with stock purchasing decisions. In this scenario, the first impulse of the investors was
to invest in a particular stock "DQ" essentially at random. Our goal is to improve hypothetical returns over this. Further
our intention was to do this efficiently with code in VBA.

##Results
	In our analysis, 2017 had across the board better stock returns than 2018. In [2018](/resources/output2018.png) few stocks in our data had positive
returns, the exceptions being [ENPH](https://www.nasdaq.com/market-activity/stocks/enph) and [RUN](https://www.nasdaq.com/market-activity/stocks/run).
Between the two, RUN had a very slightly higher yearly return of 84 to 82 percent. On the secondary objective of achieving an efficient run time,
run times reduced from about one fifth of a second to [about one tenth](/resources/2018.png) of a [second](/resources/2017.png).

##Summary
	refactoring code was advantageous in making algorithms more efficient, obtaining a 50% run time reduction in this case! It took several hours though,
and the time savings were still limited to less than a second of run time. More generally refactoring is helpful in managing memory usage and can make run times
so short as to be unnoticeable. On the other hand it takes subtantial amounts of time to refactor code.

