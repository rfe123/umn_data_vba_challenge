# umn_data_vba_challenge
Module 2 - VBA Challenge

Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock. The result should match the following image:

Moderate solution

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

Hard solution

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

NOTE
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Other Considerations
Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

SUMMARY

The solution leverages one Sub, containing several Loops. These loops are described as follows:

Loop through each Worksheet in the provided file.

  Iterate through the rows in the initial table - assuming the table is sorted by Ticker Symbol, then date.
  
    Store the row # of the first entry for a ticker symbol.
    
    When the ticker changes, summarize the data in the rows from first entry to now.
      Difference between start and end price
      Percent change between start and end price
      Sum of Total Volume column in this Range.
      Print these results with the ticker symbol in a new summary table.
      
  Iterate through the Ticker symbol results
  
    Use Sort to capture the largest % Difference profit and % Difference loss.
    Use Sort to capture the largest total Trade Volume
    Output these results to a new Summary Table

RESULTS

2018

<img width="1045" alt="image" src="https://github.com/rfe123/umn_data_vba_challenge/assets/59402267/906bae7a-6c6f-4d15-93e2-92c6ae6710b5">

2019

<img width="1048" alt="image" src="https://github.com/rfe123/umn_data_vba_challenge/assets/59402267/0a8fa87f-466f-45ad-8904-522a16d93151">

2020

<img width="1053" alt="image" src="https://github.com/rfe123/umn_data_vba_challenge/assets/59402267/162f78bd-f7d5-4d1d-a6ef-937c8a878c66">
