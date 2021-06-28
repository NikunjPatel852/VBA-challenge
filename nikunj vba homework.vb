Sub VBA Homework(): 

'Outline all variables'      
Dim Ticker_Name As String
Dim Total_Stock As Double
Dim Table As Long
Dim Pre_Table As Long
Dim Yearly_Change As Double
Dim Yearly_Open As Double
Dim Yearly_Close As Double
Dim Percent As Double

Total_Stock = 0
Table = 2
Pre_Table = 2

'looping through all workbooks'
Dim ws As Worksheet 
For Each ws in Worksheets 

'outline the last row'
LastRow = ws.cells(rows.count,1),End(x1Up).Row 

'Creating the required new coloumes'
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

''Loop through the data and get all the values'
For i = 2 To LastRow
            
Total_Stock = Total_Stock + ws.Cells(i, 7).Value
            
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Tickers for all the stocks'
Ticker_Name = ws.Cells(i, 1).Value
ws.Range("I" & Table).Value = Ticker_Name
ws.Range("L" & Table).Value = Total_Stock
Total_Stock = 0
Yearly_Open = ws.Range("C" & Pre_Table)
Yearly_Close = ws.Range("F" & i)
Yearly_Change = Yearly_Close - Yearly_Open
ws.Range("J" & Table).Value = Yearly_Change

'work out the stocks yearly_open price'         
If Yearly_Open = 0 Then
Percent = 0

Else
Yearly_Open = ws.Range("C" & Pre_Table)
Percent = Yearly_Change / Yearly_Open

End If

'Outline the number format in %'        
ws.Range("K" & Table).NumberFormat = "0.00%"
ws.Range("K" & Table).Value = Percent

'Outline the colour guid red for negitive % and green for postive %'          
If ws.Range("J" & Table).Value >= 0 Then
ws.Range("J" & Table).Interior.ColorIndex = 4
                
Else
ws.Range("J" & Table).Interior.ColorIndex = 3

End If
            
Table = Table + 1
Pre_Table = i + 1

End If

Next i

'bounes section is to create table for greatest increase, decrease and total volume 

'start by defining the needed coloumes 
ws.Cells(2,15).Value = "Greatest % Increase"
ws.Cells(3,15).Value = "Greatest % Decrease"
ws.Cells(4,15).Value = "Greatest Total Volume"

'Outline number format 
LastRow_Value = ws.Cells(Rows.Count, 11).End(xlUp).Row
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

'outline the variables 
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total As Double
Dim Value As Double

Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Total = 0


For j = 2 To LastRow_Value

'If statment outlines and exxcuties greatest_increase 
If ws.Range("K" & j).Value > Greatest_Increase Then
Greatest_Increase = ws.Range("K" & j).Value
ws.Range("Q2").Value = Greatest_Increase
ws.Range("P2").Value = ws.Range("I" & j).Value

End If

'If statment outlines and exxcuties greatest_decrease        
If ws.Range("K" & j).Value < Greatest_Decrease Then                
Greatest_Decrease = ws.Range("K" & j).Value
ws.Range("Q3").Value = Greatest_Decrease
ws.Range("P3").Value = ws.Range("I" & j).Value

End If

'If statment outlines and exxcuties greatest_total 
If ws.Range("L" & j).Value > Greatest_Total Then
Greatest_Total = ws.Range("L" & j).Value
Ws.Range("Q4").Value = Greatest_Total
ws.Range("P4").Value = ws.Range("I" & j).Value

End If

Next J 
Next ws
End Sub 
