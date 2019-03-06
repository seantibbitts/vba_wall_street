Attribute VB_Name = "week_2_homework"
'aggVol loops through a given worksheet of stock data and does the following:
'   Calculates the yearly change and yearly percent change for each stock
'   Grabs the total stock volume for the year
'   Finds the stock with the highest percent increase
'   Finds the stock with the lowest percent increase
'   Finds the stock with the highest total volume
Sub aggVol():
    Dim ws As Worksheet
    'Turn off screen updating to speed up processing
    Application.ScreenUpdating = False
    'Iterate through each worksheet in the Excel notebook
    For Each ws In Worksheets
        'Select current worksheet
        ws.Select
        
        'Declare variables
        Dim currTicker As String
        Dim currRow As Long
        Dim currOutRow As Integer
        Dim openVal As Double
        Dim closeVal As Double
        Dim openClose As Double
        Dim bgColor As Integer
        Dim percChange As Double
        Dim volume As Double
        Dim currMaxPerc As Double
        Dim currMaxPercTicker As String
        Dim currMinPerc As Double
        Dim currMinPercTicker As String
        Dim currMaxVol As Double
        Dim currMaxVolTicker As String
        
        'Create labels in spreadsheet
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatet % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'Set counters to first value
        currOutRow = 2
        currRow = 2
        
        'Iterate down spreadsheet, continuing if the first value in the current row is not empty
        Do While Not IsEmpty(Cells(currRow, 1).Value)
            openVal = Cells(currRow, 3).Value
            volume = 0
            currTicker = Cells(currRow, 1).Value
            
            'Iterate down spreadsheet, continuing if the first value in the current row is equal to the current ticker symbol
            Do While Cells(currRow, 1).Value = currTicker
                'Function was choking when pulling in new volume amount so cast as Double
                volume = volume + CDbl(Cells(currRow, 7).Value)
                closeVal = Cells(currRow, 6).Value
                currRow = currRow + 1
            Loop
            
            'Set value of yearly change
            openClose = closeVal - openVal
            
            'Set background color for yearly change--green if increase, red if decrease, no fill if zero
            If openClose > 0 Then
                bgColor = 4
            ElseIf openClose < 0 Then
                bgColor = 3
            Else
                bgColor = 0
            End If
            
            'Set cell values for yearly amounts
            Cells(currOutRow, 9).Value = currTicker
            Cells(currOutRow, 10).Value = openClose
            Cells(currOutRow, 10).Interior.ColorIndex = bgColor
            
            'Check if the original opening value is zero
            If openVal <> 0 Then
                'If not zero, get percentage change and write it to the output cell with percent formatting
                percChange = openClose / openVal
                Cells(currOutRow, 11).Value = Format(percChange, "Percent")
                'Set current max percent variable
                If IsEmpty(currMaxPerc) Then
                    currMaxPerc = percChange
                    currMaxPercTicker = currTicker
                ElseIf currMaxPerc < percChange Then
                    currMaxPerc = percChange
                    currMaxPercTicker = currTicker
                End If
                'Set current min percent variable
                If IsEmpty(currMinPerc) Then
                    currMinPerc = percChange
                    currMinPercTicker = currTicker
                ElseIf currMinPerc > percChange Then
                    currMinPerc = percChange
                    currMinPercTicker = currTicker
                End If
            End If
            'Write accumulated volume to output cell
            Cells(currOutRow, 12).Value = volume
            
            If IsNumeric(volume) Then
                'Set current max volume variable
                If IsEmpty(currMaxVol) Then
                    currMaxVol = volume
                    currMaxVolTicker = currTicker
                ElseIf currMaxVol < volume Then
                    currMaxVol = volume
                    currMaxVolTicker = currTicker
                End If
            End If
            currOutRow = currOutRow + 1
        Loop
        
        'Set cell values for overall maxes and mins
        Cells(2, 16).Value = currMaxPercTicker
        Cells(2, 17).Value = Format(currMaxPerc, "Percent")
        Cells(3, 16).Value = currMinPercTicker
        Cells(3, 17).Value = Format(currMinPerc, "Percent")
        Cells(4, 16).Value = currMaxVolTicker
        Cells(4, 17).Value = currMaxVol
        
        'Resize columns to fit contents
        Columns("A:Q").Select
        Selection.Columns.AutoFit
    
    Next
    
    'Turn screen updating back on to view results
    Application.ScreenUpdating = True
End Sub
