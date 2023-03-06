# VBA-challenge

## Overview and Purpose of the Project
The project aims to create a script that loops through all the stocks for one year and outputs the following information:
```
The ticker symbol
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
The total stock volume of the stock. The result should match the following image:
```
## VBA Code 
Sub Homework()

Dim Ticker_Label As String
Dim Percentage_change As Double
Dim Yearly_Change As Double
Dim Total_Stock_value As Double
Dim ws As Worksheet
Dim i As Integer
Dim OpenNum As Double
Dim CloseNum As Double
Dim Counter As Integer
Dim lastrow As Double

    For Each ws In Worksheets
        ws.Activate
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1").EntireColumn.Insert
        ws.Cells(1, 9) = "Ticker"
        ws.Range("J1").EntireColumn.Insert
        ws.Cells(1, 10) = "Yearly Change"
        ws.Range("K1").EntireColumn.Insert
        ws.Cells(1, 11) = "Percentage Change"
        ws.Range("L1").EntireColumn.Insert
        ws.Cells(1, 12) = "Total Stock Volume"
        
            For i = 2 To lastrow
                Ticker_Label = Cells(i, 1).Value
                    If OpenNum = 0 Then
                    OpenNum = Cells(i, 3).Value
                    End If
                
                If Cells(i + 1, 1).Value <> Ticker_Label Then
            
                Counter = Counter + 1
                Cells(Counter + 1, 9) = Ticker_Label
                          
                CloseNum = Cells(i, 6)
                Yearly_Change = CloseNum - OpenNum
                Cells(Counter + 1, 10).Value = Yearly_Change
                
                Total_Stock_value = Total_Stock_value
                Cells(Counter + 1, 12).Value = Total_Stock_value
                
                Percentage = (Yearly_Change / OpenNum)
                Cells(Counter + 1, 11).Value = Format(Percentage, "Percent")
                
                Total_Stock_value = 0
                OpenNum = 0
                
                Else
                
                Total_Stock_value = Total_Stock_value + Cells(i, 7).Value
            
                End If
             
                    If Cells(i, 10) > 0 Then
                            Cells(i, 10).Interior.ColorIndex = 3
                
                    ElseIf Cells(i, 10) < 0 Then
                            Cells(i, 10).Interior.ColorIndex = 4
                    
                    Else
                            Cells(i, 10).Interior.ColorIndex = 0
                End If
            Next i
    Next ws
    End Sub

## Results 
![Formatted Cells](https://github.com/wteklay/VBA-challenge/blob/d48a35757676204c0449e340ae052d11f8c1cfc7/Screenshot%202023-03-06%20165014.png)
# Disclaimer 
The Excel file size was too large to upload. Please use the code above. 
