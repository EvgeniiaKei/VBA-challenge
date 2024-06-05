# VBA-challenge
4 files - 1 screenshots; 2 - Separate VBA script .txt; readme file and xlx
# Module 2 Challenge
# Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyse generated stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the bonus Challenge tasks.
# Before You Begin
1. Create a new repository for this project called VBA-challenge. Do not add this assignment to an existing repository.

2. Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.
# Files
Download the following files to help you get started:
# Instructions
Create a script that loops through all the stocks for each quarter and outputs the following information:

The ticker symbol

Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

The total stock volume of the stock.   

<img width="707" alt="MulYS_Q1" src="https://github.com/EvgeniiaKei/VBA-challenge/assets/166274251/b8c8d367-3d74-4256-b438-5ae8787e081b">


# Note
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

    'Conditional formating the cells(green=positive/red=negative)
                If myws.Range("J" & table_summary_row).Value > 0 Then
                    myws.Range("J" & table_summary_row).Interior.ColorIndex = 4
                    
                Else
                    myws.Range("J" & table_summary_row).Interior.ColorIndex = 3
                End If             

# Bonus
Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

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

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.

# Other Considerations
1. Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in just a few seconds.

<img width="659" alt="Alp_test_A" src="https://github.com/EvgeniiaKei/VBA-challenge/assets/166274251/13b2d9fd-bef9-4274-a658-4b842eadae52">

# Grading

