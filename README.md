# Stocks-analysis

  ## * Overview of Project
  
  #### Our friend Steve who's just graduated with his finance degree came to us because his parents wanted to invest all their money into DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. But Steve thinks that his parents funds should be more diversified, so he wants to analyze several green energy stocks, in addition to DAQO stock.
  #### So with the Excel file he gave us, we automated tasks through an extension of Excel: Visual Basic (VBA). We prepared a workbook to helped him know how DAQO performed in 2018, and at the click of a button he can now analyze an entire dataset, but now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years.
  #### For this request we refactored our code in Visual Basic to made the VBA run faster, so Steve can work with a larger amount of data. This also means make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
  
  ## * Results
  
  #### Our objetive now was to refactored our code so it can run faster than before. So we created a ticker index variable and used it to to access the correct index across four different arrays. We created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
  
  ### Comparison between the workbooks

  #### We tried both of the macros we created to see if we accomplished the goal of making a more efficient code. We put a timer into the code and got the following results:
  
  <img width="442" alt="VBA_Challege_2018" src="https://user-images.githubusercontent.com/113747210/194675765-901a0a30-96d9-4491-a516-4579028ec420.png">

  <img width="442" alt="VBA_Challenge_2018_2" src="https://user-images.githubusercontent.com/113747210/194675781-e0d97db7-1ede-47c2-9652-bf18501a4e01.png">

  #### The fist picture is the one we created before refactoring the code and the second one can show we did made more efficient our code by running faster, so Steve can work with larger amount of data for his research and analysis of the stock market.
  
  ## * Summary
  
  ### Advantages and disadvantages of refactoring code in general
  
  #### There are some advantages of refactoring code, first we make a more efficient code, like we said before, we can take fewer steps or use less memory. A second advantage is that by refactoring a previous code we don't have to start from zero. And finally we will reach a better and more definitive code because the first one may not be the perfect one or the best, we may need a few attempts before reaching the definitive.
  
  ### Advantages or disadvantages of the original and refactored VBA script
  
  #### The advatages we saw with refactoring the original code are that we learned more, we practiced coding and overall we reduced steps. The disadvantages are that we have to do more work, it takes more time, but in the end we get the result we were looking for, that is to deliver Steve a more efficient workbook that can handle a greater amount of data than the one he was working in before.
