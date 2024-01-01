Attribute VB_Name = "Module1"

' Create a script that loops through all the stocks for one year and outputs the following information:

    'The ticker symbol

    'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

    'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

    'The total stock volume of the stock.


Sub StockSummary():

For Each ws In Worksheets

' Find the last row in the worksheet

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


' Set a variable for each stock ticker

Dim ticker_symbol As String

' Set a variable to hold Yearly Change of stock price



Dim yearly_change As Double

yearly_change = 0

' Set a variable to hold the percentage change of stock price

Dim percent_change As Double

percent_change = 0

' Set a variable to hold Total stock volume

Dim stock_volume As LongLong

stock_volume = 0

' Track the location in the Summary Row for each entry

Dim summary_row As Integer
summary_row = 2


' Set Column Headers for Script Calculations

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' Set the Percent Change summary column to proper percentage format

ws.Columns("K").NumberFormat = "0.00%"


' Loop through stock data

For i = 2 To lastrow

' Capture the opening price of the year

        Dim initial_stock_price As Boolean

            If initial_stock_price = False Then

            Dim opening_price As Double
            opening_price = ws.Cells(i, 3).Value
    
            initial_stock_price = True
    
            End If
    
        ' Checking for a change to the ticker symbol. If there is, do this
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Set the Ticker Symbol Name
        
            ticker_symbol = ws.Cells(i, 1).Value
        
        ' Calculate the Yearly Change
        
            yearly_change = ws.Cells(i, 6).Value - opening_price
                
        ' Calculate the Percent Change in price
        
            percent_change = (ws.Cells(i, 6).Value / opening_price) - 1
                
        ' Calculate the total stock volume
        
            stock_volume = stock_volume + ws.Cells(i, 7).Value
        
        ' Print the stock ticker to the summary row
        
            ws.Range("I" & summary_row).Value = ticker_symbol
        
        ' Print the yearly change to the summary row
        
            ws.Range("J" & summary_row).Value = yearly_change
                
        ' Apply Conditional formatting to change colors of yearly change based on positive or negative value
                
            If yearly_change >= 0 Then
                    
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
                    
            Else
                
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
                    
            End If
                
                
        ' Print the percent change to the summary row
        
            ws.Range("K" & summary_row).Value = percent_change
        
        ' Print the total stock volume to the summary row
        
            ws.Range("L" & summary_row).Value = stock_volume
        
        ' Increase summary row by one
        
            summary_row = summary_row + 1
        
        ' Reset the the summary row values to 0
        
            stock_volume = 0
        
        ' Reset the initial price and opening price
        
            intial_stock_price = False
            opening_price = ws.Cells(i + 1, 3).Value
                
        ' Reset the percent change
        
            percent_change = 0
        
             
    ' If this is no change to the ticker symbol, do this
    
            Else
    
    
            stock_volume = stock_volume + ws.Cells(i, 7).Value
    
    
            End If
            

         
Next i

' Set up variables for highest % increase/decrease, volume

Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As LongLong

' Find the Maxes and Mins in the summarized table to determine above variables

greatest_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
greatest_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))

' Print Row Names

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Print variable values in new table, format cells as needed

ws.Cells(2, 17).Value = greatest_increase
ws.Cells(2, 17).NumberFormat = "0.00%"

ws.Cells(3, 17).Value = greatest_decrease
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Cells(4, 17).Value = greatest_volume

' Create loop through the summarized table to pull the Ticker Symbols associated with the determined variables

For j = 2 To lastrow

    If ws.Cells(j, 11).Value = greatest_increase Then
    
    ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
    
    ElseIf ws.Cells(j, 11).Value = greatest_decrease Then
    
    ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
    
    ElseIf ws.Cells(j, 12).Value = greatest_volume Then
    
    ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
    
    End If
    
Next j

' Reset the initial stock price Boolean to false so that the correct opening price is captured
' on the next worksheet.

initial_stock_price = False

Next ws

End Sub




