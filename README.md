java c
APCO / IASC 1P01 – Fall 2024 
Assignment 3 – 
Assignment Due Date – Nov 19th, 11:59 pm uploaded to Brightspace
Assignment Late Due Date – Nov 22nd  , 11:59 pm, subject to 25% penalty
Tasks: 
•    Part A – tracking Widget Sales and Commissions
•    Part B – analyzing Stanley Cup FinalsThe goal of this assignment is to use formulas that minimizes “ hardcoding” any numbers. The goal is to  use Cell References in formula as much as possible. This will include using Absolute Cell References and Relative Cell References.
Submission: 
Please rename the file Assign 3 - .xlsx Eg – Assign 3 – Joe Smith.xlsx Ensure that you use the desktop version of Microsoft Excel (not the online version). To download the desktop version, go to myoffice.brocku.ca, sign in with your credentials, and click the Install Office button on the right-hand side. 
Part A: Widget Sales 
For Part A, you are to complete the Widget Sales Worksheet, which will look similar to Figure 1 below when completed.

Fig 1 – Widget Sales Worksheet
You will be generating the Widget Sales for Month chart and the Commission Report. The rates will be placed in the Orange Cells and used in the formulas.
The example figure shows 5 sales, but for the template that you have been provided, you will have entries for 25 sales, with the region that it was sold in. You will also be proved with the half-filled out Commission Report – with the 4 Salesperson names and the region they work in.
Your task is to fill out the following columns:
• Prior Balance – Inventory amount prior to sale
• Items Sold – Amount sold – copied over from RandomSales Worksheet
• New Balance – Inventory amount after sale
• SubTotal – Cost of Widgets (prior to tax)
• Tax – Tax amount on the SubTotal based on TaxRate
• GrandTotal – Amount customer pays (with tax)
• Commission Owed – Commission owed to salesperson (based on SubTotal and Commission Rate)
• Total Commissions – Sum total of all commissions for that salesperson in the region
• % of Total Commissions – Percentage of all commissions in company for that Salesperson. Should add up to 100%
Add total rows for each of the columns that should have totals in them. Format the cells, so that they match the formatting of the numbers in Figure 1.
You will need to use Conditional Formatting to colour code the Commission Owed column.
In order to get the Items Sold for the 25 sales and what the Widget Cost is, follow the instructions on the RandomSales Worksheet.
Part B : Stanley Cup Finals History For Part B you are to modify the worksheets in Excel to help analyze Hockey Finals since 1918. The SCF Worksheet contains the source date that you will need for all of Part B. On it, you will  find the data for the years 1918 – 2024 for the following:
Green Columns – Winning Team (Team Name, Country team plays in, Number of Games they won in Finals, Coach)
Blue Columns – Losing Team (Team Name, Country team plays in, Number of Games they won in Finals, Coach)
Orange Columns – Playoff MVP (Player Name, Team they played on)
Lime Green Columns – Scoring Leader in Playoffs (Point Total, if they played on Winning Team)

Fig 2 – SCF Worksheet
B1 – Hockey Analysis Worksheet 
For this worksheet, you are to use formulas to answer the 10 questions that are based on the SCF Worksheet. Use the formulas that we discussed during lecture. Put your formulas in the yellow boxes (remember 代 写APCO / IASC 1P01 – Fall 2024 Assignment 3R
代做程序编程语言to use cell references).
The tenth question is the most challenging … .If the Leafs and Canadiens played each other in the Finals, then that only counts as one instance, not two.
B2 – Team Championships All Time Worksheet 

Fig 3 – Team Championships All Time Worksheet
Produce a Pivot Table that looks like the following above. It will analyze the Winning Teams   against the teams they defeated in the Finals (and how many times they did it). The Teams will be filtered on Country (Can, USA).
First Level Ordering – Number of Stanley Cups won in descending order. Eg – Montreal Canadiens is at top because they have won the most cups (23).
Second Level Ordering – Number of times the losing team has been defeated in descending order. Eg – for the Montreal Canadiens, Boston is at top because the Canadiens have defeated them more in the Finals than any other team (7).
Have subtotals for each team and leave a blank line between each team.
B3 – Since 1968 Worksheet 
You are to analyze the Finals since 1968 (when the NHL initially expanded from 6 teams to 12 teams – it currently has 32).Create a pivot chart that tracks how many Stanley Cup Wins each team has since 1968, and filter  it by Country. From that, create a Pivot Chart – a 2D Bar Chart - that shows it graphically. Create a slicer based on Winning Team Country.Fig 4 – Since 1968 – Team WinsBelow the bar chart, create another Pivot Table that tracks Stanley Cup wins by Country. From the table, create a Pivot Chart – Pie Chart - that shows the percentage of wins each country has produced.Fig 5 – Since 1968 – Wins by Country
B4 – Scoring by Decade Worksheet You are to analyze the scoring in 10 year increments, starting at 1970s and going to 2020s. You  will create 6 pivot tables, and filter on the year to produce 10 values for most decades. Below is an example of 3 of the pivot tables. For the 1970s pivot table, the year would start with 197. 
Note that there are only 9 values for the 2000s, as there was no Stanley Cup awarded in 2004 and there are only 5 values for the 2020s, as we are only halfway through the decade.

Fig 6 – Scoring by Decade – 3 of the 6 pivot tables
These Grand Totals will be used as references in the Decade Total Table. You will need to calculate the yearly average for the decade. Below the table, create a Bar Chart that graphs the average yearly total for the 6 decades.
B5 –SCF Year Lookup Worksheet 
For the last Worksheet, you are to display information pertaining to the playoffs for the Year that is typed in the Yellow Box.
You will use vlookup functions to fill in the data.
Where a field requires you to lookup two pieces of information, you will need to use the Concatenate function along with the vlookup formulas.
For Canadian Teams – display the following text based on the teams that are in the Finals:
o Winning Team                                                             o Both Teams
o Losing Team                                                                 o Neither Team
The Finals box will be colour coded based on the Value found in Canadian teams. Dark Red for when the Finals that year contains 2 Canadian teams, Pink for when the Finals that year contains just 1 Canadian team, and no Colouring when the Finals contains 2 American teams.
Below demonstrates the possibilities for the 4 different statuses.

Fig 7 - SCF Year Lookup





         
加QQ：99515681  WX：codinghelp  Email: 99515681@qq.com
