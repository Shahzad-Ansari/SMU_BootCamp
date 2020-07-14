
'Shahzad Ansari
Sub stocks()
Dim sheet As Worksheet

    For Each sheet In ActiveWorkbook.Worksheets
        sheet.Activate
        
        'Instead of having to type it in per sheet manually i do it this way
        'I could see this being in issue if in a sheet these columns are already
        'Occupied but here it is okay
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'Declare and initailize variables that will be used in computation
        Dim openingPrice As Double
        Dim yearlyChange As Double
        Dim row As Long
        Dim i As Long
        Dim ticker As String
        Dim closingPrice As Double
        Dim percentageChange As Double
        Dim volume As Double
        volume = 0
        row = 2
        
        'Define the bottom row for each sheet.
        bottomRow = sheet.Cells(Rows.Count, 1).End(xlUp).row
        
        'Im not sure why i have to have this outside and increment i+1
        'in the loop, when i use just i and declare it before the calculation
        'The values dont make sense
        openingPrice = Cells(2, 3).Value
        For i = 2 To bottomRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                '''''''''''''''''''''''''''''''''
                'Set the names for all tickers
                ticker = Cells(i, 1).Value
                Range("I" & row).Value = ticker
                ''''''''''''''''''''''''''''''''''
                'Calculate Yearly Change
                closingPrice = Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                Range("J" & row).Value = yearlyChange
                '''''''''''''''''''''''''''''''''''
                'Calculate percentage change
                'Account for Div by zero and format
                If openingPrice = 0 Then
                    percentageChange = 0
                Else
                    percentageChange = yearlyChange / openingPrice
                    Range("K" & row).Value = percentageChange
                    Range("K" & row).NumberFormat = "0.00%"
                End If
                ''''''''''''''''''''''''''''''''''''
                'Calculate volume
                volume = volume + Cells(i, 7).Value
                Range("L" & row).Value = volume
                volume = 0
                ''''''''''''''''''''''''''''''''''''
                ' Incrememnt the row and the next opening price value
                openingPrice = Cells(i + 1, 3).Value
                row = row + 1
                ''''''''''''''''''''''''''''''''''''
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
        
        'Loop through each summary tables column showing yearly change and
        'Give a color to match whether the change was positive or not
        'Red = negative trend Green = possitive trend
        Dim red As Integer
        Dim green As Integer
        Dim j As Long
        red = 10
        green = 3
        bottomRowSummary = sheet.Cells(Rows.Count, 9).End(xlUp).row
        For j = 2 To bottomRowSummary
            If Range("J" & j).Value >= 0 Then
                Range("J" & j).Interior.ColorIndex = red
            Else
                Range("J" & j).Interior.ColorIndex = green
            End If
        Next j
    Next sheet
End Sub