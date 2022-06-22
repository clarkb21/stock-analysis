# Green Stocks Analysis

## Overview of Project
 - I was assigned the task to analyze stocks to help a customer make a more informed decision about which stocks their parents should purchase. 
 - I used Excel and VBA to easily create a workbook that displayed the relative information in a user friendly format. 
 - By breaking down the return for each stock, I was able to analyze stock performance to allow my customer to make the most informed decision.

### Purpose
- Using VBA in Excel, the goal was to determine stock performance for specific green stocks during the years 2017 and 2018. 
- By applying what I have learned through VBA, I was able to create an excel spreadsheet that displayed the Ticker name, total daily volume, and return for each stock that I was analyzing. 
- I also created buttons in the workbook to allow easier access for all users, not just those who understand VBA code.

![image](https://user-images.githubusercontent.com/104038813/175048808-353c4606-533c-4441-845b-01f7e494821c.png)

- Finally, I refactored the code to allow it to run quicker and more efficient using one for loop instead of multiple nested loops. 

![image](https://user-images.githubusercontent.com/104038813/175049065-64f9bac0-d67c-471a-8db2-3bb66de3d376.png)

![image](https://user-images.githubusercontent.com/104038813/175049171-a269ce9d-a9ec-4437-87c1-d9b6e01df468.png)



## Results
### Stock Performance in 2017

- Of the 12 stocks that were recorded and analyzed during this project, 11 of them returned a positive value during 2017, meaning the stock was worth more at the end of the year than it was at the beginning of the year. 
- "DQ" and "SEDG" were the two stocks that increased in value the most over the time frame. 
- "TERP" was the sole stock that lost value during the year 2017. 
- The return column is color coded to easily show the users if the stock made or lost money during the year.

![image](https://user-images.githubusercontent.com/104038813/175045752-5bada99f-e73c-4ba5-9d94-3a2c9448d0e0.png)

 
### Stock Performance in 2018
- Of the 12 stocks that were recorded and analyzed during this project, 10 of them returned a negative value during 2018, meaning those stocks lost money from the beginning of the year to the end of the year. 
- "ENPH" and "RUN" were the two stocks were gained value, showing up as green in the image below. 
- "DQ", which was the initial stock my customer was interested in, lost the most amount of money out of all 12 stocks in 2018. 
- The total daily volume shows the number of stocks that were traded during the entire year. 

![image](https://user-images.githubusercontent.com/104038813/175047059-2dfa748d-d320-4ade-ae11-6050e0c3c65f.png)


### Execution Times
- Originally, the code created in VBA for this project was designed solely for the 12 stocks that were being analyzed.
- Nested for loops were created to allow VBA to run through all the tickers and then loops through all the rows of data to find that specific ticker and return the total volume. 
- By refactoring the code, we created output arrays to allow us to only loop through the data once, making the code run much more efficiently. 
 
 #### Original Code
 
 ![image](https://user-images.githubusercontent.com/104038813/175048534-e7c03a3a-2a28-4fa2-a59f-8ac34bfa436f.png)
 
 ![image](https://user-images.githubusercontent.com/104038813/175048668-c395faeb-7c04-4eb0-9194-62b454452b08.png)



 
 #### Refactored Code
   



