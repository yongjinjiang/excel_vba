## About
   - This is the homework project for unit 2(VBA) of [Data Visualization and Analytics](https://bootcamp.umn.edu/data/landing%20full/).
   - Used VBA scripting to analyze real stock market data.
   - Finished the [hard](https://github.com/yongjinjiang/excel_vba#hard) level and realized [challenge](https://github.com/yongjinjiang/excel_vba#challenge) feature of the homework. For more details, refer to [original text](
#the-original-text-of-the-homework-assignment) of the homework.
   - Snapshot of the calculated result:
<img src="./stock_hard.jpeg" width="600" height="400">
      
   - Main tool used: Visuail Basic Application(or, in brevity, [VBA](https://docs.microsoft.com/en-us/office/vba/api/overview/language-reference)), Excel
   
## Usage
  - Download th repo, open "total_stock_volumn.xlsm", click the button "multi_year_stock_data_hard", and wait for a few seconds for the calculation to be finished and the result shown above will appear in sheet 2 of the workbook. That's all.
    


## **_The original text of the homework assignment:_** 
# Unit 2 | Assignment - The VBA of Wall Street

## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, choose your assignment from Easy, Moderate, or Hard below.

### Files

* [Test Data](02-VBA-Scripting/Homework/alphabtical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](02-VBA-Scripting/Homework/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

### Easy

* Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.

* You will also need to display the ticker symbol to coincide with the total volume.

* Your result should look as follows (note: all solution images are for 2015 data).

![easy_solution](Images/easy_solution.png)

### Moderate

* Create a script that will loop through all the stocks and take the following info.

  * Yearly change from what the stock opened the year at to what the closing price was.

  * The percent change from the what it opened the year at to what it closed.

  * The total Volume of the stock

  * Ticker symbol

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

### Hard

* Your solution will include everything from the moderate challenge.

* Your solution will also be able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

* Solution will look as follows.

![hard_solution](Images/hard_solution.png)

### CHALLENGE

* Make the appropriate adjustments to your script that will allow it to run on every worksheet just by running it once.

* This can be applied to any of the difficulties.

### Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

## Submission

* To submit please save the following in the same folder to <https://www.dropbox.com/>.

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA Scripts as separate files.

* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.

- - -

### Copyright

Coding Boot Camp © 2018. All Rights Reserved.
