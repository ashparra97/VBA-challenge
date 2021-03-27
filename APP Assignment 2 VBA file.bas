Attribute VB_Name = "FinalCode_APP"
Sub assignment()

'establish loop for each sheet within workbook
For Each ws In ActiveWorkbook.Worksheets
    'activate worksheets
    ws.Activate
    'debug to see if code actually knows what the sheets are
    'MsgBox (ws.Name)
   
     'determine variables
        Dim yearlychange As Double
        Dim closingprice As Double
        Dim openprice As Double
        Dim ticker As String
        Dim totalSV As Double
        Dim column As Integer
        Dim sum_table As Long
        Dim lastrow As Double
        Dim percentchanges As Double
        
     'assign variables a value
        column = 1
        sum_table = 2
        openprice = Cells(2, 3).Value
        totalSV = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
    'name column titles
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
       'format column and cells with percentages
        Columns("K").NumberFormat = "0.00%"
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
        
      'establish first loop
    For I = 2 To lastrow
    
    'search for when value of cells are different from the previous
        If Cells(I + 1, column).Value <> Cells(I, column).Value Then
        
            'value of each ticker
            ticker = Cells(I, column).Value
            'value of closing price
            closingprice = Cells(I, 6).Value
            'value of total stock volume
            totalSV = totalSV + Cells(I, 7)
                  
            
            'message box to tell us each different ticker
            'MsgBox (ticker)
            
            'exception for 0 in the data for openprice
                If openprice = 0 Then
                    Cells(sum_table, 9).Value = ticker
                    'put yearly change in "Yearly Change" column
                    Cells(sum_table, 10).Value = 0
                    'put percent change in "Percent Change" column
                    Cells(sum_table, 11).Value = 0
                    'put each stock total under "Total Stock Volume" column
                    Cells(sum_table, 12).Value = totalSV
                
            'put each different ticker under the "Ticker" column
                ElseIf openprice > 0 Then
                    Cells(sum_table, 9).Value = ticker
                    'put yearly change in "Yearly Change" column
                    Cells(sum_table, 10).Value = closingprice - openprice
                    'put percent change in "Percent Change" column
                    Cells(sum_table, 11).Value = (closingprice - openprice) / openprice
                    'put each stock total under "Total Stock Volume" column
                    Cells(sum_table, 12).Value = totalSV
                End If
            
            
            'tell code to start reset so it starts adding from 0
            totalSV = 0
            'tell code to add each total separately by ticker, not running total
            openprice = Cells(I + 1, 3)
            sum_table = sum_table + 1
            
        Else:
            'if the tickers are the same, AKA not different, then add them together
            totalSV = totalSV + Cells(I, 7)
        End If
    Next I

     'variables for conditional formatting
     Dim lastrow2 As Double
     lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row
     
      'second loop for conditional formatting
     For I = 2 To lastrow2
   
        'conditional formatting green and red
        If Cells(I, 10).Value >= 0 Then
            Cells(I, 10).Interior.ColorIndex = 4
        ElseIf Cells(I, 10).Value < 0 Then
            Cells(I, 10).Interior.ColorIndex = 3
        End If
        
     Next I
     
'BONUS

'determine variables
     Dim maxPC As Double
     Dim minPC As Double
     Dim maxSV As Double
     Dim maxticker As String
     Dim minticker As String
     Dim maxSVTicker As String
     
'assign values to the variables
     maxPC = 0
     minPC = 0
     maxSV = 0
    
      'establish loop
     For I = 2 To lastrow2
     
           'greatest percent decrease
            If Cells(I, 11).Value < minPC Then
            minticker = Cells(I, 9).Value
            minPC = Cells(I, 11).Value
            End If
            
            'greatest percent increase
            If Cells(I, 11).Value > maxPC Then
            maxticker = Cells(I, 9).Value
            maxPC = Cells(I, 11).Value
            End If
            
            
            'greatest total stock volume
            If Cells(I, 12).Value > maxSV Then
            maxSVTicker = Cells(I, 9).Value
            maxSV = Cells(I, 12).Value
            End If
                
            'put tickers in "ticker" cell
            Cells(2, 16).Value = maxticker
            Cells(3, 16).Value = minticker
            Cells(4, 16).Value = maxSVTicker
            
            'put values of each in second table
            Cells(2, 17).Value = maxPC
            Cells(3, 17).Value = minPC
            Cells(4, 17).Value = maxSV
            
   Next I
Next ws
End Sub
    
   


    
    


