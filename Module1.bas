Attribute VB_Name = "Module1"
Sub Master()

Dim ws_count As Integer
Dim ws_selection As Integer
Dim ws As Worksheet
Dim StartTime As Double
Dim SecondsElapsed As Double
Dim answer As Integer

'Start Timer to check how long macro runs
StartTime = Timer

Application.ScreenUpdating = False

'Get number of worksheets in this workbook
ws_count = ActiveWorkbook.Worksheets.Count

'Ask user if they need to sort their data
answer = MsgBox("Need to Sort Data? ", vbQuestion + vbYesNo + vbDefaultButton1)

If answer = vbYes Then
    'Ensure Tables are Sorted
    Call Sort_Tables(ws_count)
End If

'Set Summary Headers
For i = 1 To ws_count
    With Sheets(i)
        .Range("J1") = "ticker"
        .Range("K1") = "Yearly Change"
        .Range("L1") = "Percent Change"
        .Range("M1") = "Total Stock Volume"
        .Range("O2") = "Greatest % Increase"
        .Range("O3") = "Greatest % Decrease"
        .Range("O4") = "Greatest Total Volume"
        .Range("P1") = "Ticker"
        .Range("Q1") = "Value"
    End With
Next i

'Run Sub to populate tables for each worksheet
For Each ws In ThisWorkbook.Worksheets
    Application.DisplayStatusBar = True
    Application.StatusBar = "Working " & ws.Name & " WorkSheet..."
    Call PopulateTables(ws)
Next ws

Sheets(1).Activate

'End Timer
SecondsElapsed = Round(Timer - StartTime, 2)

Application.ScreenUpdating = True
Application.DisplayStatusBar = False

MsgBox "Successfully Completed! " & vbNewLine & "This macro ran in " & SecondsElapsed & " seconds", vbInformation

End Sub

Sub PopulateTables(ws)

Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_vol As Double
Dim max As Double
Dim min As Double
Dim maxTotal As Double
Dim maxRow As Integer
Dim minRow As Integer
Dim maxTotalRow As Integer

'Select current worksheet
ws.Select

'Fill array based on selected worksheet
arr = ws.Range("A1").CurrentRegion.Value

'Set array length minus 1
arr_length = UBound(arr) - 1

'Set initial row destination for summary tabe
row_destination = 2

'Loop through array to determine the ticker summary
For i = 2 To arr_length

    'Logic handling the 1st row of data
    If i = 2 Then
        ticker = arr(i, 1)
        open_price = arr(i, 3)
        Cells(row_destination, 10).Value = ticker
        total_stock_vol = arr(i, 7)
        
    'Logic comparing current row to next
    ElseIf arr(i, 1) = arr(i + 1, 1) Then
        'Logic checking to see if current row is 2nd to last row
        If i = arr_length Then
            close_price = arr(i + 1, 6)
            yearly_change = close_price - open_price
            
            'Check for zeroes to avoid Divide by 0 error
            If open_price = 0 Then
                percent_change = 0
            Else: percent_change = yearly_change / open_price
            
            End If
            
            'Increment total stock vol by current row plus next row
            total_stock_vol = total_stock_vol + arr(i, 7) + arr(i + 1, 7)
            
                'Populate table with yearly change, percent change, & total stock volume
                Cells(row_destination, 11).Value = yearly_change
                Cells(row_destination, 12).Value = percent_change
                Cells(row_destination, 13).Value = total_stock_vol
        
        'If row isn't 2nd to last row, then increment total stock vol with current row
        Else: total_stock_vol = total_stock_vol + arr(i, 7)
        
        End If
    
    'Logic handling when current ticker doesn't match next ticker
    ElseIf arr(i, 1) <> arr(i + 1, 1) Then
        close_price = arr(i, 6)
        yearly_change = close_price - open_price
        
        'Check for zeroes to avoid Divide by 0 error
        If open_price = 0 Then
            percent_change = 0
        Else: percent_change = yearly_change / open_price
        
        End If
            
        'Increment total stock vol with current row
        total_stock_vol = total_stock_vol + arr(i, 7)
        
        'Populate table with yearly change, percent change, & total stock volume
        Cells(row_destination, 11).Value = yearly_change
        Cells(row_destination, 12).Value = percent_change
        Cells(row_destination, 13).Value = total_stock_vol
        
        'Reset new ticker, open price, total stock vol, & row destination
        ticker = arr(i + 1, 1)
        open_price = arr(i + 1, 3)
        row_destination = row_destination + 1
        total_stock_vol = 0
            
        'Populate table with new ticker
        Cells(row_destination, 10).Value = ticker
                       
    End If

Next i

'Call functions to apply conditional formatting and format tables
Call conditional_format
Call format_tables

'Create 2nd Summary table

'Fill array based on selected worksheet
arr2 = ws.Range("J1").CurrentRegion.Value

arr2_length = UBound(arr2) - 1

max = arr2(2, 3)
min = arr2(2, 3)
maxTotal = arr2(2, 3)

maxRow = 2
minRow = 2
maxTotalRow = 2

'Find Max % Change
For i = 3 To arr2_length
    If max < arr2(i, 3) Then
        max = arr2(i, 3)
        maxRow = i
    End If
Next i

'Find Min % Change
For i = 3 To arr2_length
    If min > arr2(i, 3) Then
        min = arr2(i, 3)
        minRow = i
    End If
Next i

'Find Max Total
For i = 3 To arr2_length
    If maxTotal < arr2(i, 4) Then
        maxTotal = arr2(i, 4)
        maxTotalRow = i
    End If
Next i

Range("Q2").Value = max
Range("Q3").Value = min
Range("Q4").Value = maxTotal

Range("P2").Value = arr2(maxRow, 1)
Range("P3").Value = arr2(minRow, 1)
Range("P4").Value = arr2(maxTotalRow, 1)



End Sub


Sub conditional_format()
'
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbGreen
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"
    
    Range("M2").Select
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0,000"
    


End Sub

Sub format_tables()
'
 Range("J1").CurrentRegion.Select

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range("J1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Font.Bold = True
    Selection.EntireColumn.AutoFit
    Selection.Interior.ColorIndex = 24
    
  
    
    Range("O1:O4, P1:Q1").Select
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Selection.Font.Bold = True
    Selection.EntireColumn.AutoFit
    Selection.Interior.ColorIndex = 24
    
     Range("P2:Q4").Select
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0,000"
    Range("O:Q").Select
    Selection.EntireColumn.AutoFit
    
    
    Range("A1").Select

End Sub

Sub Clear_Cells()

Dim ws_count As Integer

ws_count = ActiveWorkbook.Worksheets.Count

'Set Summary Headers
For i = 1 To ws_count

    With Sheets(i)
        .Range("J:Q").Clear
    End With
    
Next i

End Sub

Sub Sort_Tables(ws_count)

For i = 1 To ws_count

    Range("A1").CurrentRegion.Sort _
    key1:=Range("A1"), order1:=xlAscending, _
    key2:=Range("B1"), order2:=xlAscending, Header:=xlYes
    
Next i

End Sub
