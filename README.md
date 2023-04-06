The source file was a macro-enabled excel file "Multiple_year_stock_data" that included daily data for stocks including the ticker symbol <ticker>, date <date>, daily opening price <open>, daily closing price <close>, daily highest price <high>, daily lowest price <low> and the daily volume <vol>. The excel file inlcludes sheets for multiple years: 2018, 2019 and 2020.  On each sheet, there ar 9,000 tickers and 753,000 daily rows of data.

WIth the **stockticker_analysis()** subroutine programming, new columns were created titled: Ticker, Yearly Change, Percent Change and Total Stock Volume in columns I through L.  Running the sub populates 9,000 rows of data in these columns, one row per ticker.
The VBA code loops through the daily stock data for all stocks in each worksheet and populates one row in the new columns with the yearly change, percent change and total stock volumne for the year for that stock.
Conditional formatting was added to the background cell colors of the Yearly Change column based on the numerical value of the cell.  Negative values are in red cells, positive values are in green cells and unchanged values have nobackground color.
The code continues to run in each of the next worksheets after it completes the first sheet.

With the **find_greatest_values()** subroutine programming, new columns were created titled: Greatest % INcrease, Greatest % Decrease and Greatest Total Volume in columns N through P. Running the sub populates one row of data in these columns with one ticker per column.
The VBA code loops through the output that was recently populated into columns I through L from the sub above (stockticker_analysis())
It loops through all the data in the Percent Change column comparing values to the prior value until a lowest is identified.  The ticker associated with the lowest value is populated into the column Greatest % Decrease.
It loops through all the data in the Percent Change column comparing values to the prior value until a highest is identified.  The ticker associated with the highest value is populated into the column Greatest % Increase.
It loops through all the data in the Total Stock Volume column comparing values to the prior value until a highest is identified.  The ticker associated with the highest value is populated into the column Greatest % Increase.
The code continues to run in each of the next worksheets after it completes the first sheet.
  
![image](https://user-images.githubusercontent.com/120672518/230246816-3c1334be-902d-4096-8e15-36d79ca05982.png)
![image](https://user-images.githubusercontent.com/120672518/230246872-7f8749da-cea4-40c6-aef4-5b28cb3b7649.png)
![image](https://user-images.githubusercontent.com/120672518/230246917-8e3fb7d2-d56e-4626-a7d2-6b65c73e62a6.png)

  -------------

  The VBA code is shown below.  There are 2 subs included in the code.
  
![image](https://user-images.githubusercontent.com/120672518/230247868-2483f88f-8876-4af6-a200-e97b66cda6f7.png)
![image](https://user-images.githubusercontent.com/120672518/230247936-fcfd7cf2-ac6d-43d1-97c7-ab8b13f5a925.png)
![image](https://user-images.githubusercontent.com/120672518/230247982-55fe3acd-1dae-4c4e-874c-0356ccfc171f.png)
![image](https://user-images.githubusercontent.com/120672518/230248070-a61571d3-7b29-4043-82ed-f489e8dbeb95.png)

