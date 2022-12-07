<h1> That's a Lot of Margaritas! <br> (Using Excel to track cost and revenue for a fictitious caterer's beverage operation) </h1>

#### Overview: <br>

During my tenure as a banquet captain at a large convention hotel, I was put in charge of the bars and wine service for all of the hotel's catered events. 
This was a complex operation with many moving parts and it often raised questions from my many bosses about how I ran it. 
Was my revenue on target to surpass last year and/or the goal for this year? Am I upselling our wines? Why is my cost higher this month? 
Why did I need so many bartenders last week? Why did I spend $5000 on limes? <br>

Since the systems in which we recorded our sales, purchases, labor, and inventory usage were seperate and didn't share data, I had to come up with another method 
to tie together all the details of my day-to-day operation and make that information available to my bosses in a way that was useful to them. 
Thus began my deep dive into the world of Excel! <br>

At first this project started as a simple spreadsheet to record my bartender's sales each day so that I could compare it to the Accounting Department's 
revenue reports. But over time, I started adding more details in order to preemptively answer the frequent questions from my bosses. 
How much of each sale was liquor, beer, or wine? How much did each cost? Who were the customers? How many guests attended the events? For how long? 
Had they been here before? Which ones had the greatest impact? 
Eventually, that simple spreadsheet morphed into a much larger workbook with multiple worksheets containing multiple tables, and charts, 
and complicated array formulas in many of the cells. I later added VBA code to keep the sheets properly formatted for printing. And then more VBA code to assist 
with data entry and period-end reconciling, and still more VBA code to pull data from other workbooks. <br>

I no longer work in that industry, but as this project was very important to me for so long, I still open the workbook from time to time and make some tweeks. 
I add features I wish I had thought of while I was still working in that role. I now know that this project would have been better suited for a business intelligence 
application and/or a database application, but at that time, Excel was the only tool I had access to. I have created this repository and included a demonstration 
version of this project (with fictitious entries, of course) because I think it nicely illustrates how far you can push Excel to create solutions for your data 
analysis and visualisation needs.

#### Note: <br>
This workbook was designed for users with minimal experience with Excel, therefore there are no Pivot Tables. Also, I avoided the temptation to use any of the
newer array functions that were introduced with Microsoft365 so that this workbook would work with older versions of Excel. <br> <hr>

#### Screen1: Year At A Glance Dashboard <br>
This is a "Progress Toward a Goal" type of dashboard that uses stacked columns to illustrate how my two main KPIs (Revenue and Cost Efficiency) 
compared with the previous year and to the forecast. Note the addition of lines and arrows to indicate if the KPI are moving toward or away from the goal 
and how far they would reach by the end of the year. The two speedometers below were created by layering several pie charts on top of each other. The most recent 
14 days of data were used to calculate the speeds shown on the speedometers. Also included on this dashboard is a timeline to quickly access how much time is 
left in the year to achieve the goals and also bar charts to illustrate the top 10 contributors to each of the KPIs. With this dashboard, you can quickly 
determine if you have met your goals, and if not, whether or not you have enough time and momentum to reach them. <br>

<img src="Images/Progress-At-A-Glance-Screen.jpg"> <br> <hr>

#### Screen2: Year To Date Revenue and Cost Totals <br>

<img src="Images/YTD-Revenue-Screen.jpg"> <br> <hr>

#### Screen3: Year to Date Statistics <br>

<img src="Images/Other-Statistics-Screen.jpg"> <br> <hr>

#### Group Totals Screen: <br>

<img src="Images/Group-Totals-Screen.jpg"> <br> <hr>

#### Main Data Table Screen: <br>

<img src="Images/Main-Data-Table.jpg"> <br> <hr>

#### Clicking the Category Edit Button: <br>
Several of the columns in the Main Data Table have their entries limited to the items in drop-down selection lists. To edit the items available in these 
selection lists, the user can click the "Category Edit" button. This will display a dialog box where the user can choose which selection list they would 
like to edit. Depending on which button they click next, they will be taken to the appropriate screen.

<img src="Images/Category-Edit-Button.jpg"> <br> <hr>

#### The Category Edit Screen: <br>
This is the screen where the user can edit the Billing Categories and the Product Classifications that can be used in the Main Data Table. 
Notice that all the other screens are hidden during this process. The user must click the "Done" button to return to the workbook.

<img src="Images/Category-Edit-Screen.jpg"> <br> <hr>

#### Example of Complicated Array Formulas: <br>
There are several limitations when not using a flat normalized table for your dataset. One is that you cannot use pivot tables or charts. 
Another is that you sometimes have to use very clever array formulas (previously known as CSE formulas) to slice and access the data for your calculations. 
The formula shown below calculates the number of bars in a month that were not cash bars nor was the customer charged by the drink (known as package bars). 
Notice that you can expand the formula box and use spaces and returns within your formula to make it more readable.

<img src="Images/Using-Array-Formulas.jpg"> <br><br><br> <hr>

#### Folder/File Organization:

> **"Images"** (contains images for the ReadMe file) <br>
> **"Excel Files"** (project folder) <br>
> - *"BevReport Demo.xlsm"* (Excel Macro-Enabled Workbook used for this project) <br>

	
(Please do not move, rename, delete, or alter!)
