# Stock Analysis 

## Overview of Project

### Purpose:

1. Analyse 12 different stocks data for years 2017 & 2018 and help Steve make an informed financial decision in which stocks to invest his parents money.   

2. Automate stocks analysis by refactoring an orginial VBA code to increase efficiney in run time and make the overall code more readable. 

### Results:

#### Analyzing Stocks Data:

A. Overall, stock "RUN" has performed incredibly better. Its volume has increased by almost three times and return has increased from 5.5% in 2017 to 84.0                   % in 2018, making this the better stock than any other. 

B. Although the return on stock "ENPH" has decreased significantly by 47.6% in 2018 but volume has increased by 3 times from 2017 to 2018, thus making it also a better performing stock as well. 

![](https://github.com/YuvrajT/Stock-Analysis-/blob/848ed9b846aedf135f49f65b10db86898562be95/Other%20Resources/RUN%202017.png) ![](https://github.com/YuvrajT/Stock-Analysis-/blob/a66f99657055e770836d10adc5d5f3a48d623231/Other%20Resources/RUN%202018.png)

#### Analyzing Code:

A. The orginal VBA code which was developed when working over the module 2 had a run time of 2424.496 seconds for year 2017 and 2590.555 seconds for year 2018. 

![](https://github.com/YuvrajT/Stock-Analysis-/blob/eb91386a581305c3b40783afcc7c4f56f3757220/Other%20Resources/Orginal%20Code%202017.png) ![](https://github.com/YuvrajT/Stock-Analysis-/blob/eb91386a581305c3b40783afcc7c4f56f3757220/Other%20Resources/Original%20Code%202018.png)

B. In my orginal code I did not add formating options but it still took more time to run than the refactored code for both year. 

![](https://github.com/YuvrajT/Stock-Analysis-/blob/2b40869503faa4009b2ba1ff053087085c4b68ed/Resources/VBA_Challenge_2017.png)
![](https://github.com/YuvrajT/Stock-Analysis-/blob/2b40869503faa4009b2ba1ff053087085c4b68ed/Resources/VBA_Challenge_2018.png)

C. I believe adding a ticker index and then using the same ticker index to loop over all the rows and colums in stock data for both years helped in having lesser run time compared to original code.

![](https://github.com/YuvrajT/Stock-Analysis-/blob/279e7982771260a0975c1b78e02a573eae9b09f2/Other%20Resources/Original%20Code.png) ![](https://github.com/YuvrajT/Stock-Analysis-/blob/279e7982771260a0975c1b78e02a573eae9b09f2/Other%20Resources/Refactored%20Code.png)

### Summary:

##### Advantages of refactoring the code in general:

1. Refactoring the code definitely helps make our code look more organized by applying more logic to the problem and thus removing those extra lines of code making it more concise. 

2. It also helps in debugging the code more easily which eventually results in faster processing time. Furthermore, it helps viewes to go through the code easily and understand the code much quicker. 

###### Disadvantages:

1. Refacorting code is time consuming, not all companies have the resources to assign individual specifically for these tasks. 

#### Advantages of Orginal and Refactored Script:

1. Both codes are performing the same analysis, it eye opening that one code is taking way lesser time than other. Since this was the timme first time using VBA, orginal code helped me understand the capabilites of VBA and the refactored code helped me to think differently such as, using ticke Index. 

###### Disadvantages:

1.  Definitely it was time consuming because, I had to think differently on a problem which I had already solved. The challange requirment was to use ticker Index to loop over and get the required values.  




END OF REPORT


