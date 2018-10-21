Attribute VB_Name = "Module1"
Sub StockAnalysis()
'
' StockAnalysis Macro for all worksheets
'

'
 ' '-------------------------------
    ' Declare 1D variables
    ' '-------------------------------
    Dim unique_ticker_count As Integer
    Dim row_num As Long
    Dim ticker_name As String
    Dim previous_ticker_name As String
    Dim stock_volume As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Variant
    Dim max_incr(1 To 2) As Variant
    Dim max_decr(1 To 2) As Variant
    Dim max_vol(1 To 2) As Variant
    Dim start_time As Double
    Dim seconds_elapsed As Double
    
    ' '---------------------
    ' Loop through each sheet
    ' '---------------------
    For Each ws In Worksheets
        
        ' '------------------------------------------
        ' Sort the sheet first by Column A then by Column B
        ' '------------------------------------------
        With ws.Sort
             .SortFields.Add Key:=Range("A1"), Order:=xlAscending
             .SortFields.Add Key:=Range("B1"), Order:=xlAscending
             .SetRange Columns("A:G")
             .Header = xlYes
             .Apply
        End With
        
        ' Print transaction onto transaction history
        'Sheets(format(now(), "yyyy-MM-dd hh:mm:ss")

        ' '-------------------------------
        ' Populate headers and table labels
        ' '-------------------------------
        ws.Range("I1") = "Ticker"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"

        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        ' '-------------------------------
        ' Declare a new 2D array of arbitrarily large size and four columns
        ' Column 1(index 0): ticker name. Column 2(index 1): total volume. Column3(index 2): yearly change. Column4(index 3): percent change
        ' '-------------------------------
        ReDim stock_volume_array(9999, 3) As Variant
        
        ' '-------------------------------
        ' Initialize variables to enter loop
        ' '-------------------------------
        unique_ticker_count = 0
        row_num = 2
        ticker_name = ws.Range("A" & row_num)
        stock_volume_array(0, 0) = ticker_name
        stock_volume = CDbl(ws.Range("G" & row_num).Value)
        open_price = CDbl(ws.Range("C" & row_num).Value)
        
        ' Initialize previous_ticker_name with first ticker_name
        previous_ticker_name = ticker_name
        
        ' Initialize arrays to store ticker_name and value. Column 1(index 1): ticker name. Column 2(index 2): value
        max_incr(1) = ""
        max_decr(1) = ""
        max_vol(1) = ""
        max_incr(2) = 0
        max_decr(2) = 0
        max_vol(2) = 0
        
        ' '-------------------------------
        ' Loop through rows in Column A
        ' '-------------------------------
        Do While (ticker_name <> "")
            
            If ticker_name = previous_ticker_name Then
                
                ' Add volume of row to total stock volume
                stock_volume_array(unique_ticker_count, 1) = stock_volume_array(unique_ticker_count, 1) + stock_volume
            
            Else
            
                ' '----------------------------------------------------------------------
                ' Perform closing calculations for yearly change, percent change, and summary table
                ' '----------------------------------------------------------------------
                 ' Get close price of previous row
                close_price = ws.Range("F" & row_num - 1)
                
                ' Calculate and store yearly change
                yearly_change = close_price - open_price
                
                ' Store yearly change in array
                stock_volume_array(unique_ticker_count, 2) = yearly_change
                
                ' Calculate percent change
                If open_price <> 0 Then
                    percent_change = (close_price - open_price) / open_price
                Else
                    percent_change = "n/a"
                End If
                
                ' Store percent change in array
                 stock_volume_array(unique_ticker_count, 3) = percent_change
                 
                ' Compare percent_change with current max_incr
                If percent_change > max_incr(2) And percent_change <> "n/a" Then
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
                
                ' '--------------------------
                ' Prepare variables for new ticker
                ' '--------------------------
                ' Increment unique_ticker_count
                unique_ticker_count = unique_ticker_count + 1
                
                ' Store ticker_name and stock_volume
                stock_volume_array(unique_ticker_count, 0) = ticker_name
                stock_volume_array(unique_ticker_count, 1) = stock_volume
                
                ' Store new open price
                open_price = ws.Range("C" & row_num)
                
            End If
            
            ' '--------------------------
            ' Prepare variables for new row
            ' '--------------------------
            ' Update the previous_ticker_name to current ticker_name for next loop
            previous_ticker_name = ticker_name
            
            ' Update the row_num to the next row in Column A and then the ticker_name and stock_volume
            row_num = row_num + 1
            ticker_name = ws.Range("A" & row_num) ' new ticker name
            stock_volume = CDbl(ws.Range("G" & row_num).Value)

        Loop
        
        ' '-------------------------------------------------
        ' Loop through stock_volume_array and print out the results
        ' '-------------------------------------------------
        For i = 0 To unique_ticker_count - 1
            
            ' Print ticker name
            ticker_name = stock_volume_array(i, 0)
            ws.Range("I" & i + 2) = ticker_name
            
            ' Print total volume of stock
            stock_volume = stock_volume_array(i, 1)
            ws.Range("L" & i + 2) = stock_volume
            ' Compare stock_volume with max_vol
            If stock_volume > max_vol(2) Then
                max_vol(1) = ticker_name
                max_vol(2) = stock_volume
            End If
            
            ' Print yearly change
            yearly_change = stock_volume_array(i, 2)
            ws.Range("J" & i + 2) = yearly_change
            
            ' Conditionally format positive change in green and negative change in red (Column J only)
            If yearly_change > 0 Then
                ws.Range("J" & i + 2).Interior.Color = vbGreen
            Else
                ws.Range("J" & i + 2).Interior.Color = vbRed
            End If
            
            ' Print percent change
            ws.Range("K" & i + 2) = stock_volume_array(i, 3)
        Next i
        
        ' '-----------------------------
        ' Print out summary table and format
        ' '------------------------------
        ' Greatest % increase
        ws.Range("P2") = max_incr(1)
        ws.Range("Q2") = max_incr(2)
        ws.Range("Q2").NumberFormat = "0.00%"
        ' Greatest % decrease
        ws.Range("P3") = max_decr(1)
        ws.Range("Q3") = max_decr(2)
        ws.Range("Q3").NumberFormat = "0.00%"
        ' Greatest total volume
        ws.Range("P4") = max_vol(1)
        ws.Range("Q4") = max_vol(2)
        
        ' '--------------
        ' Range formatting
        ' '--------------
        ' Column K to percent
        ws.Range("K2:K" & (2 + unique_ticker_count)).NumberFormat = "0.00%"
        ' Autofit columns
        Cells.EntireColumn.AutoFit
        
        ' '----------------------------
        ' Reset max and min variables to 0
        ' '----------------------------
        max_incr(2) = 0
        max_decr(2) = 0
        max_vol(2) = 0
        
    Next ws

End Sub
