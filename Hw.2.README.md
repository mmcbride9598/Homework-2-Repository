# Homework-2-Repository
Repository for Hw#2
Sub HW_2()

Dim Tickler As String

Dim Summary_Table_Row As Integer

Dim Last_Row As Integer

Dim Year_change As Double

Dim Percent_Change As Double

Dim Total_Stock_Volume As Long

Dim Open_Date As Double

Dim Close_Date As Double


Summary_Table_Row = 2


For Each ws In Worksheets


For i = 2 To 79771

Tickler = ws.Cells(i, 1).Value

Open_Date = ws.Cells(i, 3).Value

Close_Date = ws.Cells(i, 5).Value

Year_change = Open_Date - Close_Date

Percent_Change = ((Close_Date - Open_Date) / Open_Date) * 100

Total_Stock_Volume = ws.Cells(i, 7).Value


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

ws.Cells(i, 1).Value = Tickler

ws.Range("J" & Summary_Table_Row).Value = Tickler

ws.Range("K" & Summary_Table_Row).Value = Year_change

ws.Range("L" & Summary_Table_Row).Value = Percent_Change

ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume

Summary_Table_Row = Summary_Table_Row + 1

Year_change = 0

Percent_Change = 0

Total_Stock_Volume = 0

Else

Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

End If

Next i

For i = 2 To 1000

If ws.Cells(i, 11) >= 0 Then

ws.Cells(i, 11).Interior.ColorIndex = 4

Else

ws.Cells(i, 11).Interior.ColorIndex = 3

End If

Next i

Next ws

End Sub
