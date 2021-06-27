Attribute VB_Name = "Module1"
Sub VBA_Challenges()


'Variables to use on all worksheets
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

'Number of rows
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set up Ticker to compare with
LastTicker = ""

'Sum of the stock volume
vol = 0


'Loop through each worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Column titles of the tables on each worksheet
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

'Set up Max, Min, Count to record the greatest ticker
Max = 0
MaxIndex = 0
Min = 0
MinIndex = 0
Count = 0
CountIndex = 0

'Counter that will keep track of rows that will write to new table and do it on all the ws
Counter = 1

    'Loop through each row till a row after the end of list
    For i = 2 To (LastRow + 1)
    
        'If same ticker name then add stock volume to vol
        If Cells(i, 1).Value = LastTicker Then
            
            
            vol = vol + Cells(i, 7)
            
        'If it is the very first ticker:
        ElseIf Counter = 1 Then
            
          
            Counter = Counter + 1
            
            'Record ticker symbol and write it to the table just below the header
            LastTicker = Cells(i, 1).Value
            Cells(Counter, 9).Value = LastTicker
            
            'Record OpenInitial for calculating Change
            OpenInitial = Cells(i, 3).Value
        
        'If the ticker is different:
        Else
            
            'Increment the counter
            Counter = Counter + 1
            
            'Record new ticker symbol and write it to the table
            LastTicker = Cells(i, 1).Value
            Cells(Counter, 9).Value = LastTicker
            
            'Calculate change and record in Change
            CloseFinal = Cells(i - 1, 6).Value
            Change = CloseFinal - OpenInitial
                
                'Use a conditional to avoid divide by zero error
                'If OpenInitial is zero, then set percent to 0
                If OpenInitial = 0 Then
                    
                    Percent = 0
                
                'Else write to Percent
                Else
                
                    Percent = Change / OpenInitial
                                            
                End If
            
            'Use the Percent to compare to the Max, Min and Sum to Count to find the greatest throughout the worksheet
            If Percent > Max Then
                
                Max = Percent
                MaxIndex = Counter
                
            ElseIf Percent < Min Then
                
                Min = Percent
                MinIndex = Counter
                
            End If
            
            If vol > Count Then
                
                Count = vol
                CountIndex = Counter
                
            End If
                            
            'Write Change and Percent to the table
            Percent = FormatPercent(Percent, 2)
            Cells(Counter - 1, 10) = Change
            Cells(Counter - 1, 11) = Percent
                
            'If Change is positive, then make it green
            If Cells(Counter - 1, 10).Value > 0 Then
            
                Cells(Counter - 1, 10).Interior.ColorIndex = 4
            
            'If Change is negative, then make it red
            ElseIf Cells(Counter - 1, 10).Value < 0 Then
            
                Cells(Counter - 1, 10).Interior.ColorIndex = 3
                
            End If
                            
            'Record new OpenInitial
            OpenInitial = Cells(i, 3).Value
            
            'Write Sum to the table and reset it
            Cells(Counter - 1, 12).Value = vol
            vol = 0
            
        End If
        
    Next i
    
    'Greatest value on the 2nd table
    Range("O2").Value = Cells(MaxIndex - 1, 9).Value
    Range("O3").Value = Cells(MinIndex - 1, 9).Value
    Range("O4").Value = Cells(CountIndex - 1, 9).Value
    Range("P2").Value = FormatPercent(Max, 2)
    Range("P3").Value = FormatPercent(Min, 2)
    Range("P4").Value = Count
    
Next

starting_ws.Activate



End Sub

