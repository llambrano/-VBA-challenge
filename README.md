# -VBA-challenge
by Libardo Lambrano

> ## Table of Contents

* Intro 
* Submission 



> ### Intro

> In this excercise I used VBA scripting to analyze stock market data. The script grabs the ticker name from a database of daily quotes, it calculates yearly changes (net and percentage) as well as total volume. A summary table extracts the stock with the biggest percentage increase, the one with the lowest percentage decrease and the stock with the largest transacted volume. 

> ### Submission 

This is the original database, it has six rows: `<ticker>`, `<date>`, `<open>`, `<high>`, `<low>`, `<close>`, `<vol>`. Tickers are organized organized by year (one per sheet). Here is an exsample of `2016` with 797K records.

![](//VBAStocks/images/01_original_table.png)

**Step 1**

* Collect all tickers in a new column `I`, yearly changes in column `J`, percentage changes in column `K` and total stock volume in column `L`. 

* Conditional formating applied in column `J` to highlight positive and negative variances. 


![](/VBAStocks/images/01_step1.png)

**Step 2**

* Append summary table including the stock with the largest and lowest percentage variance within a year, and the one with highest trading volume. 

![](/VBAStocks/images/01_step2.png)









