# StockAnalyzerWB
C# Excel Workbook Add-in that creates new worksheets that show a table of stock option prices and the potential profits and ROI for various scenarios.

It pulls in the option data for a given stock for four months- the coming 2 months and for the January options for the next 2 years.
If you have ThinkOrSwim running (provided by TD Ameritrade) then the option prices will be updated in real time.


This now requires TD Ameritrade API's, because Yahoo finance stopped providing stock tables. 
So to use this you will need to register as a developer at TD Ameritrade.


How to Install:
Register to create a Developer Account at https://developer.tdameritrade.com/ .
Register an app there to get your "OAuth User ID".
Edit  "publish\Application Files\StockAnalyzerWB_1_1_0_1\AuthData.txt.deploy" in notepad, and put your "OAuth User ID" on the first line of the file (by itself).
run "publish\setup.exe"