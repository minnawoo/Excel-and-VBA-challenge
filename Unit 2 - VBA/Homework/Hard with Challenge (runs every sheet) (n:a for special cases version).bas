Attribute VB_Name = "Module1"
Sub StockAnalysis()
'
' StockAnalysis Macro for all worksheets
' Calls functions below
' 1 = Easy
' 1 & 2 = Moderate
' 1 & 2 & 3 = Hard

'
' 1. Calculate and print total volume of each ticker name for all worksheets
Call TotalVolumeOfEachStock

' 2. Calculate the yearly change and percent change from year open price to year close price
Call CalculateYearlyChangeOfEachStock

' 3. Calculate the "Greatest % Increase", "Greatest % Decrease" and "Greatest total volume"
Call CreateSummaryTable

End Sub
Function CreateSummaryTable()
'
' CalculateYearlyChangeOfEachStock Function for all worksheets
'

'
    ' Declare a counter variable for number of worksheets to loop through
    Dim num_worksheets As Integer
    num_worksheets = ActiveWorkbook.Worksheets.Count
    
    ' '---------------------
    ' Loop through each sheet
    ' '---------------------
    For sheet_num = 1 To num_worksheets
        
        ' Populate table labels and headers
        Sheets(sheet_num).Range("O2") = "Greatest % Increase"
        Sheets(sheet_num).Range("O3") = "Greatest % Decrease"
        Sheets(sheet_num).Range("O4") = "Greatest Total Volume"
        Sheets(sheet_num).Range("P1") = "Ticker"
        Sheets(sheet_num).Range("Q1") = "Value"
        
        ' Declare and initialize row_num with 2
        Dim ticker_num As Long
        ticker_num = 2
        
        ' Declare and initialize ticker_name with name of first stock in Column I
        Dim ticker_name As String
        ticker_name = Sheets(sheet_num).Range("I" & ticker_num)
        
        '  Declare and initialize arrays to store ticker_name and value. Column 1(index 1): ticker name. Column 2(index 2): value
        Dim max_incr(1 To 2) As Variant
        Dim max_decr(1 To 2) As Variant
        Dim max_vol(1 To 2) As Variant
        max_incr(1) = ""
        max_decr(1) = ""
        max_vol(1) = ""
        max_incr(2) = 0
        max_decr(2) = 0
        max_vol(2) = 0
        
        ' Declare and initialize percent_change and total_volume variables with first row of data in Column I
        Dim percent_change As Double
        Dim stock_volume As Double
        percent_change = Sheets(sheet_num).Range("K" & ticker_num)
        stock_volume = Sheets(sheet_num).Range("L" & ticker_num)
                
        ' Loop through Column I ticker names
        Do While (ticker_name <> "")
            
            ' Compare percent_change with current max_incr
            If percent_change > max_incr(2) Then
                ' Update the current max_incr
                max_incr(1) = ticker_name
                max_incr(2) = percent_change
            End If
            
            ' Compare percent_change with current max_decr
            If percent_change < max_decr(2) Then
                 ' Update the current max_decr
                max_decr(1) = ticker_name
                max_decr(2) = percent_change
            End If
            
            ' Compare stock_volume with max_vol
            If stock_volume > max_vol(2) Then
                ' Update the current max_vol
                max_vol(1) = ticker_name
                max_vol(2) = stock_volume
            End If
            
            ' Update ticker_num, ticker_name, percent_change, and stock_volume for next row
            ticker_num = ticker_num + 1
            ticker_name = Sheets(sheet_num).Range("I" & ticker_num)
            If Sheets(sheet_num).Range("K" & ticker_num) = "n/a" Then ' n/a percent change
                Exit Do
            End If
            percent_change = Sheets(sheet_num).Range("K" & ticker_num)
            stock_volume = Sheets(sheet_num).Range("L" & ticker_num)
        Loop
        
        ' Print out the results
        ' Greatest % increase
        Sheets(sheet_num).Range("P2") = max_incr(1)
        Sheets(sheet_num).Range("Q2") = max_incr(2)
        Sheets(sheet_num).Range("Q2").NumberFormat = "0.00%"
        ' Greatest % decrease
        Sheets(sheet_num).Range("P3") = max_decr(1)
        Sheets(sheet_num).Range("Q3") = max_decr(2)
        Sheets(sheet_num).Range("Q3").NumberFormat = "0.00%"
        ' Greatest total volume
        Sheets(sheet_num).Range("P4") = max_vol(1)
        Sheets(sheet_num).Range("Q4") = max_vol(2)
        
        ' Autofit columns
        Cells.EntireColumn.AutoFit
        
    Next sheet_num
End Function
Function CalculateYearlyChangeOfEachStock()
'
' CalculateYearlyChangeOfEachStock Macro for all worksheets
'

'
    ' Declare a counter variable for number of worksheets to loop through
    Dim num_worksheets As Integer
    num_worksheets = ActiveWorkbook.Worksheets.Count
    
    ' '---------------------
    ' Loop through each sheet
    ' '---------------------
    For sheet_num = 1 To num_worksheets
        
        ' Populate headers
        Sheets(sheet_num).Range("J1") = "Yearly Change"
        Sheets(sheet_num).Range("K1") = "Percent Change"
        
        ' Declare and initialize row_num with 2
        Dim ticker_num As Long
        ticker_num = 2
        
        ' Declare and initialize ticker_name with name of first stock in Column I
        Dim ticker_name As String
        ticker_name = Sheets(sheet_num).Range("I" & ticker_num)
        
        ' NOTE: THIS STRATEGY IS BASED ON THE ASSUMPTION THAT COLUMN A WILL ALWAYS HAVE ALPHABETICAL LISTING
        '  AND THAT COLUMN B WILL ALWAYS HAVE CHRONOGICAL LISTING BASED ON THE TICKER NAME IN COLUMN A.
        ' IF THIS ASSUMPTION BECAME FALSE, PLEASE ADD CODE TO SORT THE SHEETS FIRST SO THAT COMPUTATION IS MORE EFFICIENT/
        ' Declare and initialize row_num with 2
        Dim row_num As Long
        row_num = 2
        
        ' Declare open and close price variables and percent change variables to be used in inner loops
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        
        'Loop through Column I ticker names
        Do While (ticker_name <> "")
            
            ' Loop through rows in Column A and find the first and last instance of each ticker_name
            Do While (found = False)
                
                ' Check if ticker name in Column A matches ticker_name from Column I
                If Sheets(sheet_num).Range("A" & row_num) = ticker_name Then
                    
                    ' Store open price
                    open_price = Sheets(sheet_num).Range("C" & row_num)
                    
                    ' Continue looping through Column A until ticker_name is different (and take the previous row's close price)
                    Do While (found = False)
                    
                        ' Increment row_num
                        row_num = row_num + 1
                        
                        ' Check if ticker name in Column A stops matching the ticker_name from Column I
                        If Sheets(sheet_num).Range("A" & row_num) <> ticker_name Then
                            
                            ' Store close price of previous row
                            close_price = Sheets(sheet_num).Range("F" & row_num - 1)
                            
                            ' Calculate yearly change
                            yearly_change = close_price - open_price
                            
                            ' Print yearly change in Column J
                            Sheets(sheet_num).Range("J" & ticker_num) = yearly_change
                            
                            ' Conditionally format positive change in green and negative change in red (Column J only)
                            If yearly_change > 0 Then
                                Sheets(sheet_num).Range("J" & ticker_num).Interior.Color = vbGreen
                            Else
                                Sheets(sheet_num).Range("J" & ticker_num).Interior.Color = vbRed
                            End If
                            
                            ' Calculate percent change
                            If open_price = 0 Then
                                percent_change = 9999 ' CHANGE THIS TO YOUR PREFERENCE"
                            Else
                                percent_change = (close_price - open_price) / open_price
                            End If
                            
                            ' Print percent change in Column K with formatting
                            If percent_change = 9999 Then
                                Sheets(sheet_num).Range("K" & ticker_num) = "n/a"
                            Else
                                Sheets(sheet_num).Range("K" & ticker_num) = percent_change
                            End If
                            Sheets(sheet_num).Range("K" & ticker_num).NumberFormat = "0.00%"
                            
                            ' Set found to True
                            found = True
                        End If
                    Loop
                End If
            Loop
            
            ' Update ticker_num and ticker_name for next row
            ticker_num = ticker_num + 1
            ticker_name = Sheets(sheet_num).Range("I" & ticker_num)
            
            ' Reset found to false for next ticker_name
            found = False
        Loop
    Next sheet_num
End Function
Function TotalVolumeOfEachStock()
'
' TotalVolumeOfEachStock Function for all worksheets
'

'
    ' Declare a counter variable for number of worksheets to loop through
    Dim num_worksheets As Integer
    num_worksheets = ActiveWorkbook.Worksheets.Count
    
    ' Loop through each sheet
    For sheet_num = 1 To num_worksheets
        
        ' Sort
        With Sheets(sheet_num).Sort
             .SortFields.Add Key:=Range("A1"), Order:=xlAscending
             .SortFields.Add Key:=Range("B1"), Order:=xlAscending
             .SetRange Columns("A:G")
             .Header = xlYes
             .Apply
        End With
    
        ' Populate headers
        Sheets(sheet_num).Range("I1") = "Ticker"
        Sheets(sheet_num).Range("L1") = "Total Stock Volume"
        
        ' Declare and initialize unique_ticker_count with 0
        Dim unique_ticker_count As Integer
        unique_ticker_count = 0
        
        ' Declare and initialize row_num with 2
        Dim row_num As Long
        row_num = 2
        
        ' Declare and initialize ticker_name with name of first stock and stock_volume with volume of first stock
        Dim ticker_name As String
        Dim previous_ticker_name As String
        Dim stock_volume As Double
        ticker_name = Sheets(sheet_num).Range("A" & row_num)
        previous_ticker_name = ticker_name
        
        stock_volume = CDbl(Sheets(sheet_num).Range("G" & row_num).Value)
        
        ' Declare and initialize boolean, ticker_exists, with False
        Dim found As Boolean
        ticker_exists = False
        
        ' Declare a new stock_volume_array of size unique_ticker_count and two columns
        ReDim stock_volume_array(9999, 1) As Variant ' Column 1(index 0): ticker name. Column 2(index 1): total volume
        
        ' Loop through rows and find the number of unique ticker names
        Do While (ticker_name <> "")
            
            If ticker_name = previous_ticker_name Then
                
                ' Add volume of row to total stock volume
                stock_volume_array(unique_ticker_count, 1) = stock_volume_array(unique_ticker_count, 1) + stock_volume
            Else
            
                ' Increment unique_ticker_count
                unique_ticker_count = unique_ticker_count + 1
                
                ' Store ticker_name and stock_volume
                stock_volume_array(unique_ticker_count, 0) = ticker_name
                stock_volume_array(unique_ticker_count, 1) = stock_volume
            End If
            
            ' Update the previous_ticker_name to current ticker_name for next loop, the row_num to the next row
            ' in Column A, and the corresponding ticker_name and stock_volume
            previous_ticker_name = ticker_name
            row_num = row_num + 1
            ticker_name = Sheets(sheet_num).Range("A" & row_num) ' new ticker name
            stock_volume = CDbl(Sheets(sheet_num).Range("G" & row_num).Value)
            
            ' Reset ticker_exists to false
            ticker_exists = False
        Loop
        
        ' Loop through array and print out results
        For i = 0 To UBound(stock_volume_array)
            ticker_name = stock_volume_array(i, 0)
            If ticker_name = "" Then
            
                ' Nothing left to check in the rest of the array, so exit for loop
                Exit For
            Else
            
                ' Print ticker name and total volume of each stock
                Sheets(sheet_num).Range("I" & i + 2) = ticker_name
                Sheets(sheet_num).Range("L" & i + 2) = stock_volume_array(i, 1) ' total volume of stock
            End If
        Next i
    Next sheet_num
End Function

