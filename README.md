# Stock Analysis

## Overview of Project

  Our friend, Steve has given us an Excel file containing the stock data he wants us to analyze. We used this data to create a Macro enabling Steve to analize the entire dataset with a click of a button. Steve now wants to expand his data set but in order to adapt to a large data pull our VBA code could be refactored to loop through the data faster. In this challange we will be comparing our run time from the initial time completed in this Module:
  ```green_stocks```
 We will then compare it to our refactored code in:
   ```VBA_Challenge.xlsm```
From here we can easily compare any improvments to our VBA time which would provide the adaptability of the code needed to compensate a larger data set and compare any improvments or challanges to the code along the way. 
## Results
**Initial Times**

```green_stocks_2017```

![VBA_Challenge_2017_Time_1](https://user-images.githubusercontent.com/115853964/199649916-35687074-9ad9-4574-b50f-bf6ce042d932.png)

```green_stocks_2018```

![VBA_Challenge_2018_Time_1](https://user-images.githubusercontent.com/115853964/199649941-188769d8-9ee4-403a-a5e6-c977386f39af.png)

The initial times for ```green_stocks```VBA run time takes longer than we would like for Steve's plan to add more data to the set, refactoring the code can allow us to make the necessary changes to adapt our code to run in a more efficient and timely way. This will allow the code to use less memory, improve the logic for the code to make it easier for Steve to use analysing larger pools of Stock data. 

**Refactoring Code**

![VBA_Challenge_Code_1](https://user-images.githubusercontent.com/115853964/199650275-0ccd0181-139f-4886-81c3-aee65b0c52cd.png)

Creating the ```TickerIndex``` to reflect the intial set = 0. From here our code will begin to loop it over all our data 


![VBA_Challenge_Code_2](https://user-images.githubusercontent.com/115853964/199650320-24ea2185-fb70-4b33-8f65-541b52950081.png)

Creating three output arrays; ```tickerVolumes``` are *Long*: Positive and negative whole numbers between –2,147,483,648 and 2,147,483,647, stored in 32 bits and ```tickerStartingPrices``` and  ```YearsVlickerEndingPrices``` as *Single*: Outlined by -3.402823E38 to -1.401298E-45 for negative values 1.401298E-45 to 3.402823E38 for positive values.


![VBA_Challenge_Code_3](https://user-images.githubusercontent.com/115853964/199650338-fe2b41b4-e492-46cc-aca3-0a40747ad467.png)

Increasing the ```tickerIndex``` to reflect the change in ```i```as it processes through the loops we establised if the next row’s ticker doesn’t match the previous row’s ticker.

**Improved VBA Time**

```VBA_Challenge.xlsm_2017```

![VBA_Challenge_2017_Time_2](https://user-images.githubusercontent.com/115853964/199650486-fb43ee88-c959-4b3e-9b49-dc4beff92a9f.png)

```VBA_Challenge.xlsm_2018```

![VBA_Challenge_2018_Time_2](https://user-images.githubusercontent.com/115853964/199650506-95808446-0ffa-4528-bbf0-3cc0f16b4463.png)

Run time of VBA is impoved with refactoring of code. 

### Summary 
-What are the advantages or disadvantages of refactoring code?

Refactoring code can be very beneficial to the overall functionability and easy of use of code for a specific task. In this case running a larger pool of data for Steve that is of a time sensitive nature such as stocks it is important to save a tool that is running at a capacity that is not only capable of handling a larger data pull but offering a more user friendly and translatable platform to preform the task. 
In this case the increase in runtime was a huge benefit to Steve's want to expand data, however it did not come without challenges. The largest challenge I faced in restructurings was working backwards and having the tendency to *get lost in the syntax* for lack of better terms. This along with errors like two calls to  ```yearsValue ``` doubled my already drastically improved run time. This *Small* error caused an additional 45 seconds that could change a lot in certain data sets.

-How do these pros and cons apply to refactoring the original VBA script?


The Pros to adding to the original VBA script is you are able to adapt the source and this can be very beneficial if you are working with a shorter set of code, however on the opposite end, with more intricate strings of code there is more room for error and changing the original VBA code changes any references back to the original code to make further adaptations

