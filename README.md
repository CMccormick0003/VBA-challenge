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
  
  ![image](https://user-images.githubusercontent.com/120672518/230245941-439d7b45-f53e-4fc8-841b-bf32004d578e.png)

  ![image](https://user-images.githubusercontent.com/120672518/230246010-a031d483-2c83-48fe-80e6-6ffa75978782.png)

  ![image](https://user-images.githubusercontent.com/120672518/230246082-813b069e-1e47-4108-b53c-be3a005a2adc.png)
  -------------

  The VBA code is shown below.  There are 2 subs included in the code.
  
Sub stockticker_analysis()

Dim stocktotal As Double
Dim rowindex As Long
Dim change As Double 'Where change is the change in stock price
Dim tablerow As Integer
Dim start As Long
Dim rowcount As Long
Dim percentchange As Double
Dim sheet As Worksheet

For Each sheet In Worksheets 'Loop through each worksheet in the workbook

tablerow = 0
stocktotal = 0
change = 0
start = 2
dailychange = 0

'Create all the labels for the data on the table
sheet.Range("I1").Value = "Ticker"
sheet.Range("j1").Value = "Yearly Change"
sheet.Range("k1").Value = "Percent Change"
sheet.Range("l1").Value = "Total Stock Volume"
sheet.Range("I1").Font.Bold = True
sheet.Range("j1").Font.Bold = True
sheet.Range("k1").Font.Bold = True
sheet.Range("l1").Font.Bold = True

'Calculate the row count for each worksheet - confirmed with msgbox
rowcount = sheet.Cells(Rows.Count, 1).End(xlUp).Row

'Create a loop that goes through each row

For rowindex = 2 To rowcount

    If sheet.Cells(rowindex + 1, 1).Value <> sheet.Cells(rowindex, 1).Value Then
    
        stocktotal = stocktotal + sheet.Cells(rowindex, 7).Value
        
        If stocktotal = 0 Then
        
'Print the results
            sheet.Range("I" & 2 + tablerow).Value = Cells(rowindex, 1).Value
            sheet.Range("J" & 2 + tablerow).Value = 0
            sheet.Range("K" & 2 + tablerow).Value = "%" And 0
            sheet.Range("L" & 2 + tablerow).Value = 0
        Else

'Identify the cell with the open stock price to calculate the change in stock price
            If sheet.Cells(start, 3) = 0 Then
                For find_value = start To rowindex
                    If sheet.Cells(find_value, 3).Value <> 0 Then
                    start = find_value  'establishes the row index with the open price
                    Exit For
                    End If
                Next find_value
            End If
        
'Calculate the stock change and percentage change
            change = (sheet.Cells(rowindex, 6) - sheet.Cells(start, 3))
            percentchange = change / sheet.Cells(start, 3)
            
            start = rowindex + 1
            
'Print the final calculations to the table
            sheet.Range("I" & 2 + tablerow) = sheet.Cells(rowindex, 1).Value
            sheet.Range("j" & 2 + tablerow) = change
            sheet.Range("j" & 2 + tablerow).NumberFormat = "0.00"
            sheet.Range("K" & 2 + tablerow).Value = percentchange
            sheet.Range("K" & 2 + tablerow).NumberFormat = "0.00%"
            sheet.Range("L" & 2 + tablerow).Value = stocktotal
            
'Conditional color formatting for the columns with change in stock price
            Select Case change
                Case Is > 0
                    sheet.Range("J" & 2 + tablerow).Interior.ColorIndex = 4
                Case Is < 0
                    sheet.Range("J" & 2 + tablerow).Interior.ColorIndex = 3
                Case Else
                    sheet.Range("J" & 2 + tablerow).Interior.ColorIndex = 0
            End Select
            
            
        
        End If
        
'Reset data before the loop runs again
    stocktotal = 0
    change = 0
    tablerow = tablerow + 1
    
    Else
    
'Sum the stock total volume for the table
'The variable is stored in the IF statement
    stocktotal = stocktotal + sheet.Cells(rowindex, 7).Value
    
    End If

Next rowindex

Next sheet
End Sub


Sub find_greatest_values()

    Dim sheet As Worksheet
    Dim last_row As Long
    Dim max_increase_ticker As String
    Dim max_increase_value As Double
    Dim max_decrease_ticker As String
    Dim max_decrease_value As Double
    Dim max_volume_ticker As String
    Dim max_volume_value As Double
    
' Loop through each worksheet
    For Each sheet In ThisWorkbook.Worksheets
        
' Add labels and headers for the new columns and rows to report the stocks with the greatest changes
        sheet.Cells(2, 15).Value = "Greatest % Increase"
        sheet.Cells(3, 15).Value = "Greatest % Decrease"
        sheet.Cells(4, 15).Value = "Greatest Total Volume"
        sheet.Cells(1, 16).Value = "Ticker"
        sheet.Cells(1, 17).Value = "Value"
        sheet.Range("O2").Font.Bold = True
        sheet.Range("O3").Font.Bold = True
        sheet.Range("O4").Font.Bold = True
        sheet.Range("P1").Font.Bold = True
        sheet.Range("Q1").Font.Bold = True
        
' Get the last row number in the worksheet
        last_row = sheet.Cells(Rows.Count, "I").End(xlUp).Row
        
' Reset the variables for the new worksheet
        max_increase_ticker = ""
        max_increase_value = 0
        max_decrease_ticker = ""
        max_decrease_value = 0
        max_volume_ticker = ""
        max_volume_value = 0
        
' Loop through each row in the worksheet to find the greatest % increase, greatest % decrease, and greatest total volume
        For i = 2 To last_row

' Get the ticker, yearly change, percent change, and total volume values
            Dim ticker As String
            Dim yearly_change As Double
            Dim percent_change As Double
            Dim total_volume As Double
            ticker = sheet.Cells(i, "I").Value
            yearly_change = sheet.Cells(i, "J").Value
            percent_change = sheet.Cells(i, "K").Value
            total_volume = sheet.Cells(i, "L").Value
            
' Check for the greatest % increase
            If percent_change > max_increase_value Then
                max_increase_value = percent_change
                max_increase_ticker = ticker
            End If
            
' Check for the greatest % decrease
            If percent_change < max_decrease_value Then
                max_decrease_value = percent_change
                max_decrease_ticker = ticker
            End If
            
' Check for the greatest total volume
            If total_volume > max_volume_value Then
                max_volume_value = total_volume
                max_volume_ticker = ticker
            End If
        Next i
        
' Print the results for the worksheet
        sheet.Cells(2, 16).Value = max_increase_ticker
        sheet.Cells(2, 17).Value = max_increase_value
        sheet.Cells(3, 16).Value = max_decrease_ticker
        sheet.Cells(3, 17).Value = max_decrease_value
        sheet.Cells(4, 16).Value = max_volume_ticker
        sheet.Cells(4, 17).Value = max_volume_value
    
'Run the code for the next worksheet
    Next sheet

End Sub
