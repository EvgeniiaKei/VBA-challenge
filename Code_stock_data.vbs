Sub Quarterly_Analysis()

'Challenge 2 - Kozodeeva Evgeniia
'Declare Worksheet as "mws" each worksheet
Dim myws As Worksheet
For Each myws In ThisWorkbook.Worksheets
myws.Activate

'Define all variables
'(Total Stock Volum = total_St_Vol, Quarterly_change = quar_change,great total volum = great_tot_vol)
Dim ticker As String
Dim open_price As Double
Dim closing_price As Double
Dim percentage_change As Double
Dim total_St_Vol As Double
total_St_Vol = 0
Dim quar_change As Double
Dim great_tot_vol As Double
great_tot_vol = 0
'Other variables and values for loop to start
Dim PreviousStockPrice As Long
PreviousStockPrice = 2
Dim table_summary_row As Integer
table_summary_row = 2
Dim greatest_increase As Double
greatest_increase = 0
Dim greatest_decrease As Double
greatest_decrease = 0

'Label Column Headers and Tables
myws.Range("P1").Value = "Ticker"
myws.Range("Q1").Value = "Value"
myws.Range("O2").Value = "Greatest % Increase"
myws.Range("O3").Value = "Greatest % Decrease"
myws.Range("O4").Value = "Greatest Total Volume"
myws.Range("I1").Value = "Ticker"
myws.Range("J1").Value = "Quarterly Change"
myws.Range("K1").Value = "Percent Change"
myws.Range("L1").Value = "Total Stock Volume"

'For each stock find the Quaterly Change, Percent Change, and Total Stock Volume

'value of the last row for column A
Lastrow = myws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through first rows for stock
For i = 2 To Lastrow

        'Find the value of the Total Stock Volume
        total_St_Vol = total_St_Vol + myws.Cells(i, 7).Value
        
        'Record for change in stock ticker in the summary table with ticker name and total_St_Vol and reset total_St_Vol back to 0
        If myws.Cells(i + 1, 1).Value <> myws.Cells(i, 1).Value Then
        
                ticker = myws.Cells(i, 1).Value
                myws.Range("I" & table_summary_row).Value = ticker
                myws.Range("L" & table_summary_row).Value = total_St_Vol
                
                total_St_Vol = 0
                
                'Open and close price, quaterly change, and percentage change
                open_price = myws.Range("C" & PreviousStockPrice)
                close_price = myws.Range("F" & i)
                quar_change = close_price - open_price
                myws.Range("J" & table_summary_row).Value = quar_change
                
                'statement to determine percent change with If
                If open_price = 0 Then
                
                    percentage_change = 0
                    
                Else
                    open_pice = myws.Range("C" & PreviousStockPrice)
                    percentage_change = quar_change / open_price
                    
                End If
                
                'Percentage change in summary table using the % format
                myws.Range("K" & table_summary_row).Value = percentage_change
                myws.Range("K" & table_summary_row).NumberFormat = "0.00%"
                
                'Conditional formating the cells(green=positive/red=negative)
                If myws.Range("J" & table_summary_row).Value > 0 Then
                    myws.Range("J" & table_summary_row).Interior.ColorIndex = 4
                    
                Else
                    myws.Range("J" & table_summary_row).Interior.ColorIndex = 3
                End If
                
                If myws.Range("J" & table_summary_row).Value = 0 Then
                    myws.Range("J" & table_summary_row).Interior.ColorIndex = 0
                    
                    
                End If
                
                'Initiate task to go to next row
                table_summary_row = table_summary_row + 1
                PreviousStockPrice = i + 1
                
            End If
            
            Next i

'The last row for column K
Lastrow = myws.Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To Lastrow

    'First determine the Greatest Total Volume
    If myws.Range("L" & i).Value > great_tot_vol Then
       great_tot_vol = myws.Range("L" & i).Value
       myws.Range("Q4").Value = great_tot_vol
       myws.Range("P4").Value = myws.Range("I" & i).Value
       
    End If
    'Second Greatest % Increase
    If myws.Range("K" & i).Value > greatest_increase Then
        greatest_increase = myws.Range("K" & i).Value
        myws.Range("Q2").Value = greatest_increase
        myws.Range("P2").Value = myws.Range("I" & i).Value
        
    End If
    'Third Greatest % Decrease
    If myws.Range("K" & i).Value < greatest_decrease Then
        greatest_decrease = myws.Range("K" & i).Value
        myws.Range("Q3").Value = greatest_decrease
        myws.Range("P3").Value = myws.Range("I" & i).Value
        
    End If
    
    'Change format to "%" for Greatest % Increase and Decrease
    myws.Range("Q2").NumberFormat = "0.00%"
    myws.Range("Q3").NumberFormat = "0.00%"
    
Next i

Next myws

        
End Sub