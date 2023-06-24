**Week 2 challenge.**

I have submited 3 files for your consideration:
- a Macro-enabled Excel file called "alphabetical_testing"
- a Macro-enabled Excel file called "Multiple_year_stock_data"
- a .vbs file called "Sub stocksummary()"
- a README file.
- a screenshot: Screenshot 2023-06-23 192240
- a screenshot: Screenshot 2023-06-23 192343

The .vbs file contains the code that I have run in both Excel files.  The code is also included in the excel files as a macro.

I created the code using the following sources of documentation:
- AskBCS Learning Assistants: to help with 2 issues:
	-**Issue 1**: for the summary table 1, I could not figure out why my code was not adding all the values for each ticker.  The last 	one was not being included in the total sum of the volumes.  To solve the issue, this was the recommendation I received:"I do spot 	the reason why the volume is missing the last value. In your code, the totalstockvolume is added at the end of the for loop. In 	other words, it first checks if the current stock and next stock are the same. If they are not the same, then the summary is created 	and values are reset. Then it adds the volume to totalstockvolume. Since the cell is written before before the last value is added 	to totalstockvolume, then the last value is not going to be included to the total.  So yup, that line of code should be placed to 	the start of the loop, not the end of the loop. And in your code, the reset only occurs when the if statement is triggered, so it 	will be not be reset when the current stock is the same as the next stock, but will reset when the stocks are not the same."
	
	-**Issue 2**: for the second summary table, to assign the value of the tickers to the maximum percent increase, decrease, and value.  	I could not figure out how we could bring the ticker name.  I tried conditionals and loops, but cannot make it work.  This was the 	recommendation I received: "When setting the conditional for ws.Range("Q2") and ws.Cells(summaryrow, "I") you are comparing more 	than just the item inside the cells. You will also be comparing the Row, Column, Value, etc. Since ws.Range("Q2") is not in the same 	row/column as ws.Cells(summaryrow, "I"), this conditional will always return false. The only comparison that should be made is the 	Value within that cell, hence the conditional should be set as: `If ws.Range("q2").Value = ws.Cells(summaryrow, "I").Value Then`"