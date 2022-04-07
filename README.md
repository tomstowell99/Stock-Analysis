# Automated Stock Analysis for Steve

## Overview of Project
I will be reviewing the twelve green energy stocks for the years 2017 and 2018. We be refining our code from prior work, “refactoring” to make the update process more efficient and thereby faster. 

### Purpose
The purpose of this new code is to show that by refining your code. You can create code that does not have as many loops or steps. This will cause the code to run faster. 
 
## Analysis and Challenges



### Analysis of original code
The original code I had used two loops to perform the analysis. One nested within the other. The first loop would iterate over the stock code and the inner loop would then iterate over each row in the spread sheet.
This approach would have each row evaluated 12 times, once for each stock. With this approach. The code took .734 and .746 seconds to evaluate years 2017 and 2018 respectively.

### Analysis of New Code
The new code eliminated the need of the inner loop. We created data arrays where we could store data specific to that stock as we evaluated each row. This eliminated the need to loop through the entire data set every time to evaluate each new stock. This shortened are code cycle time to .117 and .113 for years 2017 and 2018 respectively.


## Summary

### Advantages of refactoring
The refactoring allowed us to reduce the run time of the code significantly from .7 seconds to .1. This was accomplished through adding the Array variables. With the array variables set up. We could store data specific to that stock versus waiting until the end of our loop and printing out.

### Disadvantage of refactoring
A disadvantage to refactoring was the code was a little more challenging to follow.

### Conclusion speed of code improvement
By eliminating the need for the second for loop. We no longer were revisiting our data multiple times. This is what reduced the run time on our VBA code.
I believe this will be an important concept for clean programming going forward as we work with larger data sets.


