# stock-analysis

## Project Overview

### Purpose

The purpose of this week’s assignment was to explore the capabilities of Visual Basic for Applications (VBA) to analyze Excel data.

### Data Analyzed

The specific dataset I analyzed included stark market data for 12 different stocks in the years 2017 and 2018. Throughout our weekly modules, we created a code that would create a table of stock tickers and their associated Total Daily Volumes and Returns.  For this challenge, we were asked to ‘refactor’ the starter code to make it run faster and more efficiently.  

### Deliverables

The deliverable for this assignment was a properly refactored code that executed an ‘AllStockAnalysis’ macro.  The code was to include an input box for the year to be analyzed (2017 or 2018) as well as a message box that displayed the length of time it took to run the code.  

## Results

As you can see from the screenshots below, my refactored code ran the code for ‘2017’ in 0.164 seconds.  It ran my refactored code for ‘2018’ in 0.148 seconds.  The results also reveal that, in general, 2017 had better returns for most of the identified stocks than 2018.  

2017 Refactored Code

 ![image](https://user-images.githubusercontent.com/92705556/147314682-79acd2a8-b99f-478a-8932-cf420fb44846.png)


2018 Refactored Code

 ![image](https://user-images.githubusercontent.com/92705556/147314691-799944b9-ebd6-4f55-8b42-de16ef8b24a9.png)
 
When comparing the time taken to run the refactored code to the time taken to run the original code (shown below), you can see that the refactored code ran 2017 data in one-six the time and the 2018 in one-seventh the time.  This confirms that the refactored code is faster and more efficient than the original code.  

2017 Original Code

![image](https://user-images.githubusercontent.com/92705556/147314708-28383d7b-b460-4001-98c1-7c8b0373910e.png)

2018 Original Code

 ![image](https://user-images.githubusercontent.com/92705556/147314715-30fd6d0b-14fc-4b3b-9f0c-ca02f9eae9ad.png)


### Differences Between Codes

The main difference between the original code and the refactored code is that the refactored code uses more arrays and a tickerIndex variable.  The Original Code does not use these arrays and instead uses nested for statements.  

Examples of Refactored Code with arrays and variables

![image](https://user-images.githubusercontent.com/92705556/147314746-4de8797a-bc8d-432d-aa59-6e07d56350c5.png)

![image](https://user-images.githubusercontent.com/92705556/147314762-906c52b2-3090-4562-b660-b02cc9e3a87e.png)

 
Examples of Original Code

 ![image](https://user-images.githubusercontent.com/92705556/147314774-5fd9a1df-7507-40ae-80ad-b03cb34cd132.png)

![image](https://user-images.githubusercontent.com/92705556/147314786-021da4e1-5b38-4533-8ed2-1b043cc63103.png)
       
## Summary

### Advantages of Refactoring In General

Refactoring is advantageous because it simplifies and cleans up original code that make be riddled with lengthy and complicated steps.  By cleaning up the original code, it makes it macro run faster, can reduce bugs, and is easier to understand.  These benefits are amplified using arrays.  

### Disadvantages of Refactoring In General

Though refactoring can be beneficial, it also has disadvantages.  First of all, you must first have original code to work with.  If you don’t have original code, you cannot refactor it. Refactoring also takes time and can be difficult if you are not familiar with VBA code and functions.  While refactoring may result in faster run times, the ultimate solution/result it the same as the original output.  Thus, if you are not worried about run times, refactoring might not be worth your time.  

### Advantages and Disadvantages of refactoring the provided stocks analysis script

Overall, I believe the refactored code is cleaner than the original code, and it obviously runs faster. That being said, as a beginning coder not very familiar with code and/or VBA, I believe some of the technicalities with arrays and variables can be confusing and cause errors in the code.  Refactoring may be better suited for more experienced coders that better understand why, how, and where errors occur. I definitely need a lot more practice before I can even remotely call myself ‘experienced.’
