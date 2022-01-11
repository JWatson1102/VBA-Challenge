Attribute VB_Name = "Module1"
Sub stocks()

Dim totalVol As Double
Dim openVal As Double
Dim closeVal As Double
Dim tick As String
Dim nextTick As String
Dim row As Long
Dim cont As Boolean
Dim firstTick As Boolean
Dim printRow As Integer

row = 2
printRow = 2
firstTick = True
cont = True



Do While cont = True

    'stores the current ticker value as well as the next row
    tick = Cells(row, 1).Value
    nextTick = Cells(row + 1, 1).Value

    'sums the total volume as the loop runs through the sheet
    totalVol = totalVol + Cells(row, 7).Value

    'checks to see if row is the open date for the ticker
    If firstTick = True Then
        openVal = Cells(row, 3).Value
        firstTick = False
    End If

    'checks to see if next line is a different ticker
    'also confirms the close date for the ticker
 If (nextTick <> tick) Then
        closeVal = Cells(row, 6).Value

        'now that the data has been stored, we will "print" it to excel
        
        'prints the ticker value
        Cells(printRow, 9).Value = tick
        'prints the change stock value
        Cells(printRow, 10).Value = closeVal - openVal
        
        'formats the cell colors during the loop to save time
        If (Cells(printRow, 10).Value < 0) Then
            Cells(printRow, 10).Interior.ColorIndex = 3
            
        ElseIf (Cells(printRow, 10).Value > 0) Then
            Cells(printRow, 10).Interior.ColorIndex = 4
            
        Else
            Cells(printRow, 10).Interior.ColorIndex = 6
        
        End If
        
        'calculates the % change in stock Value and controls for dividing by zero
        If (openVal <> 0) Then
            Cells(printRow, 11).Value = (Cells(printRow, 10).Value / openVal) * 100
            
        Else
            Cells(printRow, 11).Value = "N/A"
            
        End If
        
        Cells(printRow, 12).Value = totalVol

        'resets total values to prepare for next ticker
        totalVol = 0
        printRow = printRow + 1
        firstTick = True



    End If

    'should end the loop once it reads that the next line is empty
    If (nextTick = "") Then
        cont = False
    End If

    'adds to row to move the loop to the next line
    row = row + 1


Loop

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Change"
Cells(1, 11).Value = "% Change"
Cells(1, 12).Value = "Total Volume"




End Sub



