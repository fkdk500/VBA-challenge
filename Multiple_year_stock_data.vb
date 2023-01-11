
'Steps:
' ----------------------------------------------------------------------------
' Part I:

'1.loop through each year of stock data and grab the total amount of volume each stock had over the year.

'2.Display the ticker symbol in new column to accord to the total volume.

'3 add the openprice,closeprice and the total stcok volume in the respective column

'Part II:
'1.Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" In new summary sheet

Sub Multiple_year_stock_data()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' Variable declaration of Ticker, openprice, closeprice and stcok volume to Calculate Yearly price change,
        ' Percent Change and total stock volume
        ' --------------------------------------------
        Dim Open_price As Double
        Dim Close_price As Double
        'To exceed larger value of long data type
        'Refernce:https://sodocumentation.net/vba/topic/3418/data-types-and-limits
        Dim stockvolume As Variant
        Dim ticker As String
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_yearly_price As Double
        Dim percent As Double
        Dim ticker_row As Integer
              
      ' Loop through each row to Keep track of each tickers in the summary table
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Price_Row = 2
        stockvolume = 0
        yearly_change = 0
        percent_change = 0
                
       ' Create a Variable to Hold File Name and Last Row
        Dim WorksheetName As String

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
                   
        'add the word Ticker, Year Change, Persent Change and Total Volume as Cloumn Head
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Year Change"
        ws.Cells(1, 11).Value = "Persent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Add conditional and looping through each row
        
        For i = 2 To LastRow
        openprice = ws.Cells(Summary_Table_Row, 3).Value
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
               ticker = ws.Cells(i, 1).Value
            
            'Calculate yearly change and percent change
               Open_price = ws.Range("C" & Price_Row).Value
               Close_price = ws.Range("F" & i).Value
               yearly_change = Close_price - Open_price
               
            If Open_price = 0 Then
               percent_change = 0
            Else
                percent_change = yearly_change / Open_price
            End If
                            
              'Print the ticker,yearchange and stockvolume in the Summary Table
               
              'add to the ticker
              ws.Range("I" & Summary_Table_Row).Value = ticker
              
              ' Add to the year change
              ws.Range("J" & Summary_Table_Row).Value = yearly_change
              
              ' Add to the percent change
              ws.Range("K" & Summary_Table_Row).Value = percent_change
              
              ws.Range("K" & Summary_Table_Row).Style = "Percent"
              ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
              
              'Add to the stockvolume
              ws.Range("L" & Summary_Table_Row).Value = stockvolume
                 
               
              'Add one to the summary table row
               Summary_Table_Row = Summary_Table_Row + 1
               Price_Row = i + 1
            
              ' Reset the stockvolume
               stockvolume = 0
          
            Else
            
            stockvolume = stockvolume + ws.Cells(i, 7).Value
            'reference:https://stackoverflow.com/questions/2202869/what-does-the-on-error-resume-next-statement-do
            'when you encounter an error just continue at the next line.
            On Error Resume Next
            End If
                      
        Next i

        ' --------------------------------------------
        ' greatest total volume, max & min
        ' --------------------------------------------

       'Declare variables for cell formatting
        Dim yearLastRow As Long
        yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Add Loop for cell formatting
        For i = 2 To yearLastRow
        
        'Add Conditional for cell formatting
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
            
        'Declare variables for finding max & min
        Dim percentLastRow As Long
        percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        Dim percent_max As Double
        percent_max = 0
        Dim percent_min As Double
        percent_min = 0
        
        
        'Add Loop for finding max & min
        For i = 2 To percentLastRow
        
        'Add Conditional for max & min(Greatest increase and Greatest decrease)
            If percent_max < ws.Cells(i, 11).Value Then
                percent_max = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = percent_max
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 15).Value = "Greatest % Increase"
                
            ElseIf percent_min > ws.Cells(i, 11).Value Then
            
                percent_min = ws.Cells(i, 11).Value
                ws.Cells(3, 17).Value = percent_min
                ws.Cells(3, 17).NumberFormat = "0.00%"
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 15).Value = "Greatest % Decrease"
                
            End If
        Next i
        
        
        'Declare variable for greatest total volume
        Dim totalVolumeRow As Long
        totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
        Dim totalVolumeMax As Double
        totalVolumeMax = 0
        
        'Add Loop for finding greatest total volume
        For i = 2 To totalVolumeRow
        
        'Add Conditional for greatest total volume
            If totalVolumeMax < ws.Cells(i, 12).Value Then
                totalVolumeMax = ws.Cells(i, 12).Value
                ws.Cells(4, 17).Value = totalVolumeMax
                'Refernce: https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba
                ws.Cells(4, 17).NumberFormat = "0.00E+00"
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 15).Value = "Greatest Total Volume"
                
                'Column width fixing
                'Reference:https://stackoverflow.com/questions/22322550/vba-change-excel-cell-width
                ws.Range("O1").ColumnWidth = 20
                ws.Range("J1").ColumnWidth = 12
                ws.Range("K1").ColumnWidth = 12
                ws.Range("L1").ColumnWidth = 12
        
        
            End If
        Next i
        
    Next ws
        
        

End Sub
