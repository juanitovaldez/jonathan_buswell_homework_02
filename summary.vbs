Attribute VB_Name = "Module1"

Option Explicit
'Filters the list into a column of unique tickers'

Sub CreateUniqueList()
Dim lastrow As Long
Dim lastunique As Long


'sets up headers'
Cells(1, 9).Value = "<ticker>"
Cells(1, 10).Value = "<total_vol>"
Cells(1, 11).Value = "<delta_year>"
Cells(1, 12).Value = "<%dy>"

Cells(1, 15).Value = "<ticker>"
Cells(1, 16).Value = "<value>"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

'gets the index of the last populated row'
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
   'filters out unique values and copies them to a summary table with dimension 1xlastrow
    ActiveSheet.Range("A2:A" & lastrow).AdvancedFilter _
    Action:=xlFilterCopy, _
    CopyToRange:=ActiveSheet.Range("I2"), _
    Unique:=True
'get the index of the last unique ticker

lastunique = Cells(Rows.Count, "I").End(xlUp).Row
'Sum total values for each ticker and write them into column
'collect all relevant stats
'total_volume = sum of each unique tickers daily volumns
'delta_year = max.date(open) - max.date(close)
'% change = (max.date(close)-min.date(open))/min.date(open)

'iterate through each unique ticker
'coolect variables in accumulators
'find open and close dates
Dim raw_ticker_range As Range, raw_ticker As Range
Dim uni_ticker_range As Range, uni_ticker As Range
Dim max_date As Long, min_date As Long
Dim total_vol As Double, dy As Double, p_dy As Double
Dim year_open As Double, year_close As Double
Dim max_vol As Double, max_dy As Double, min_dy As Double

Dim max_vtic As String, max_dtic As String, min_dtic As String

Set uni_ticker_range = Range("I2:I" & lastunique)
Set raw_ticker_range = Range("A2:A" & lastrow)

For Each uni_ticker In uni_ticker_range
    ' re initialize variables for next ticker
    max_date = 0
    min_date = 99999999
    total_vol = 0
    year_open = 0
    year_close = 0
    For Each raw_ticker In raw_ticker_range
        If uni_ticker.Value = raw_ticker.Value Then
            
            'sum items in the volume column for each unique ticker symbol
            total_vol = total_vol + raw_ticker.Offset(0, 6).Value
            
            'check for max and min dates gets year open/close values
            If max_date < raw_ticker.Offset(0, 1).Value Then
                max_date = raw_ticker.Offset(0, 1).Value
                year_close = raw_ticker.Offset(0, 5)
            End If
            If min_date > raw_ticker.Offset(0, 1).Value Then
                min_date = raw_ticker.Offset(0, 1).Value
                year_open = raw_ticker.Offset(0, 2).Value
            End If
            
        End If
            
    Next raw_ticker
'    Debug.Print uni_ticker
'    Debug.Print max_date
'    Debug.Print min_date
'    Debug.Print year_open
'    Debug.Print year_close
    dy = year_close - year_open
'   Don't divide by zero
    If year_open = 0 Then
        'fudge factor
        year_open = 0.01
        p_dy = (year_close - year_open) / year_open
    Else
        p_dy = dy / year_open
    End If
    
'   populate table in excel sheet with our stats
    uni_ticker.Offset(0, 1).Value = total_vol
    uni_ticker.Offset(0, 2).Value = dy
    uni_ticker.Offset(0, 3).Value = p_dy
Next uni_ticker

' This section will find the maximum volume, delta and min delta using the same patter to find the closing dates
For Each uni_ticker In uni_ticker_range
    If max_vol < uni_ticker.Offset(0, 1).Value Then
        max_vol = uni_ticker.Offset(0, 1).Value
        max_vtic = uni_ticker.Value
    End If
     If max_dy < uni_ticker.Offset(0, 3).Value Then
        max_dy = uni_ticker.Offset(0, 3).Value
        max_dtic = uni_ticker.Value
    End If
     If min_dy > uni_ticker.Offset(0, 3).Value Then
        min_dy = uni_ticker.Offset(0, 3).Value
        min_dtic = uni_ticker.Value
    End If
Next uni_ticker
Cells(1, 15).Value = "<ticker>"
Cells(1, 16).Value = "<value>"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

Cells(2, 15).Value = max_dtic
Cells(3, 15).Value = min_dtic
Cells(4, 15).Value = max_vtic

Cells(2, 16).Value = max_dy
Cells(3, 16).Value = min_dy
Cells(4, 16).Value = max_vol

End Sub

Sub summarize_sheets()

Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
            ws.Activate
            Call CreateUniqueList
         Next ws

End Sub





