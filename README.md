# VBA Stock Analysis Challenge

## Overview of Project

In this project our goal was to analyze stock data with a view to providing good, evidence-based investment advice to Steve's parents. We started with an Excel spreadsheet full of stock data, and over the course of the lessons we developed a VBA script that produced an analysis of stocks according to their volume and return (first for one stock of special interest to Steve's parents, then for multiple stocks at once). That code, although effective for its purpose, was not as efficient as it could have been (in terms of its runtime especially), so the specific goal of this project was to refactor the existing code for efficiency.

## Results

We were able to refactor the code successfully using arrays as a tool for organizing our data. The initial subroutine we built during the lessons, AllStocksAnalysis(), used a variable "ticker" to hold the name of the ticker we were working on in the analysis:

<img width="423" alt="image" src="https://user-images.githubusercontent.com/99628051/156950272-582efa25-1e83-4ee9-8bbd-77ab6768eaa2.png">

Coupled with a variable "totalVolume" to hold the volume of the current ticker, as well as variables "endingPrice" and "startingPrice" to hold the ending & starting prices of each index, the initial subroutine was able to loop through the different tickers and produce the desired output for the analysis:

<img width="515" alt="image" src="https://user-images.githubusercontent.com/99628051/156950396-dbbca95a-1a73-482d-93a3-302b02b318b4.png">

While this code functioned successfully without much issue, it ran for each year with runtimes on the order of half a second to a full second, as evidenced by the following screenshots:

<img width="419" alt="image" src="https://user-images.githubusercontent.com/99628051/156950540-153531cc-9989-464a-b3b7-75ee59c5b2a3.png">

<img width="418" alt="image" src="https://user-images.githubusercontent.com/99628051/156950572-16197646-3b31-457e-a208-2ee461f5c28a.png">

In order to make our code more time-efficient, we used arrays to organize the ticker names, total volume, and starting + ending prices in the refactored version of the macro, called AllStocksAnalysisRefactored().

The screenshot below shows the declaration of a tickerIndex variable and the arrays, along with a for loop that initializes all the array values to 0:

<img width="470" alt="image" src="https://user-images.githubusercontent.com/99628051/156950697-569ddb5f-e071-49d9-aa3d-592b94c0a49e.png">

With these arrays in place, we were able to restructure our main loop with the principal goal of populating each array, with no need to continuously re/overwrite placeholder variables each time. This was done with the same conditional structure as before:

<img width="692" alt="image" src="https://user-images.githubusercontent.com/99628051/156950845-00a86a38-989e-4d72-af70-3592828de4b7.png">

After that loop completed, we were then able to produce output by looping through our arrays instead of having to output the contents of the placeholder variables before they were overwritten:

<img width="599" alt="image" src="https://user-images.githubusercontent.com/99628051/156950929-9f0bb360-4681-439d-a549-d6c603d8f45b.png">

Although there was more memory required to store these values in arrays, the improvement in time efficiency speaks for itself:

<img width="424" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/99628051/156950962-57422b9d-49ee-4235-b881-6de9a5e12fcf.png">

<img width="423" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/99628051/156950973-3e92c678-3b2c-4696-a312-b555d2116ba7.png">

## Summary

### What are the advantages or disadvantages of refactoring code?

In general, there are many advantages of refactoring code: like any problem-solving process, code writing is unlikely to result in a totally optimal solution the first time, so iterating the solution process by consistently refactoring, reevaluating, revisiting, and rewriting code is a good strategy for ensuring different parts of a complex project are coded as optimally as possible. The resulting gains in efficiency may seem small in the short run -- for example, in our analysis a change from .75 seconds to less than a quarter of a second for the 2017 analysis might not feel like a huge leap -- but when datasets and the algorithms we use to analyze them become more complex, these increases in efficiency could reduce computation time or memory use by several orders of magnitude, in some cases. Moreover, continually revising code gives one the opportunity to freshly ask the question of how well the code serves its intended purpose, and if there are any innovations that could be brought to bear on the code that would help it to fit more harmoniously in the context of a larger project.

Code refactoring is not entirely without disadvantages, however. Refactored code can make tradeoffs in the name of efficiency -- for example, an algorithm can be made more time-efficient by increasing the demands it makes on memory. In addition, refactored code can be broken, such that a function or program under revision can no longer serve its initial purpose until the refactored code is fully debugged. Although good version control via Git or some other platform can usually control the worst fallout from this downside -- you can always go back to the un-refactored, or original, code, and make your edits in a separate branch to as to fiddle and troubleshoot without disturbing the integrity of the whole project -- it can still be a major headache to have to painstakingly debug new code just to get to the same standard of functionality you had attained previously!

### How do these pros and cons apply to refactoring the original VBA script?

The pros mentioned above bore out very well in the case of our VBA script. As mentioned above, we were able to greatly increase the time efficiency of our subroutine, which could make a major difference in runtime if the algorithm were tasked with running through a very large dataset with terabytes of stock data. In addition, revisiting the subroutine led this programmer to get to ask some questions that he otherwise may have treated as "asked and answered" because the subroutine was working -- for example, instead of hard-coding the tickers that we wanted to analyze, why not create another subroutine that accepts user input of which tickers the user is interested in analyzing? This would produce a whole new set of coding challenges -- how to accept an open number of inputs, how to make an array big enough to store them all, how to validate the input to ensure what the user is typing in are actually valid tickers, etc. -- but the resulting code would greatly increase the flexibility of our analytical tool for Steve and for his parents.

Similarly, the disadvantages of refactoring code shone through here too -- when I accidentally made some index mistakes while writing the refactored subroutine, as well as mistakenly (or, to be a tad more charitable to myself, unnecessarily) trying to use the tickerIndex variable to output the results, it was very frustrating to have bugs get in the way of what had been fully functioning code just minutes ago! In addition to my purely subjective response to the refactoring, however, there is also a much more objective consequence: the refactored subroutine requires more space in memory to function properly, since the work that had previously been done by furiously writing and outputting and rewriting three placeholder variables is now done by populating three arrays. Just as our new leaps in time efficiency could pay dividends (pun gently intended) for larger data sets, the new memory requirements could prove a liability in future, more complex analyses.
