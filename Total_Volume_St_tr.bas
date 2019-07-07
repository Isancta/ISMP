Attribute VB_Name = "Module1"
Sub Total_Stock_VOlume()

' Create Variables to hold outputs


Dim Tickers As String
Dim Total_Volume As Double
Total_Volume = 0
Dim Summary_Table As Integer
Summary_Table = 2

' Create Variable inputs

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all Tickers

For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Tickers symbol

Tickers = Cells(i, 1).Value
Total_Volume = Total_Volume + Cells(i, 7).Value

' Print Outputs in Summary_Table

Range("I" & Summary_Table).Value = Tickers
Range("J" & Summary_Table).Value = Total_Volume

' Add 1 to Summary-Table

Summary_Table = Summary_Table + 1

'Reset Total_Volume
Total_Volume = 0

Else



Total_Volume = Total_Volume + Cells(i, 7).Value

End If

 Next i

Cells(1, 9).Value = "Tickers"

Cells(1, 10).Value = "Total Volume"

End Sub

