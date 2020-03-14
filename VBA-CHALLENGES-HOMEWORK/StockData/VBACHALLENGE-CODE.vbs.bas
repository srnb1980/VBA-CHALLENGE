'Uncomment the below line to import the module into visual basic developer
'Attribute VB_Name = "Module1"
Sub VBACHALLENGE()
'DECLARING ALL THE NEEDED VARIABLES
Dim StartValue As Double
Dim FinalValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim GTPercentIncrease As Double
Dim GTPercentDecrease As Double
Dim GTStockvolume As Double
Dim TotalStockVolume As Double
Dim LastRow As Integer
Dim SummaryRow As Integer
Dim WS_Count As Integer
'SETTING THE WORKSHEET COUNT
WS_Count = ActiveWorkbook.Worksheets.Count
'SETTING THE FOR LOOP TO GO THROUGH EVERY SHEET
For J = 1 To WS_Count
'SETTING THE ACTIVE WORKSHEET
'MsgBox (ActiveWorkbook.Worksheets(J).Name)
 Set Sht = ActiveWorkbook.Worksheets(J)
'LastRow = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row
'MsgBox (Sht.Cells(Sht.Rows.Count, 1).End(xlUp).Row)
'SETTING THE HEADER TITLES FOR THE SUMMARY RECORDS
  Sht.Cells(1, 9) = "Ticker"
  Sht.Cells(1, 10) = "Yearly Change"
  Sht.Cells(1, 11) = "Percent Change"
  Sht.Cells(1, 12) = "TotalStockVolume"
  Sht.Cells(1, 15) = "Ticker"
  Sht.Cells(1, 16) = "Value"
  Sht.Cells(2, 14) = "Greatest % Increase"
  Sht.Cells(3, 14) = "Greatest % Decrease"
  Sht.Cells(4, 14) = "Greatest Total Volume"
'SETTING THE INITIAL VALUES BEFORE PROCESSING EACH SHEET
SummaryRow = 2
TotalStockVolume = 0
StartValue = Sht.Cells(2, 3).Value
'MsgBox StartValue
'SETTING THE FOR LOOP TO GO OVER EACH ROW
For I = 2 To Sht.Cells(Sht.Rows.Count, 1).End(xlUp).Row
'IF ELSE CLAUSE TO CHECK IF THE CURRENT TICKER IS DIFFERENT FROM NEXT TICKER
 If (Sht.Cells(I, 1).Value = Sht.Cells(I + 1, 1).Value) Then
  FinalValue = Sht.Cells(I, 6).Value
  TotalStockVolume = TotalStockVolume + Sht.Cells(I, 7)
 Else
'SETTING UP THE NEEDED VALUES IN THE SUMMARY
  FinalValue = Sht.Cells(I, 6).Value
  YearlyChange = FinalValue - StartValue
'SPECIAL HANDLING IF STARTVALUE IS ZERO
  If (StartValue = 0) Then
   PercentChange = YearlyChange / 1
  Else
   PercentChange = YearlyChange / StartValue
  End If
  TotalStockVolume = TotalStockVolume + Sht.Cells(I, 7)
  Sht.Cells(SummaryRow, 9) = Sht.Cells(I, 1).Value
  Sht.Cells(SummaryRow, 10) = YearlyChange
'SETTING THE COLORINDEX BASED ON INCREASE OR DECREASE
  If (YearlyChange <= 0) Then
   Sht.Cells(SummaryRow, 10).Interior.ColorIndex = 3
  Else
   Sht.Cells(SummaryRow, 10).Interior.ColorIndex = 4
  End If
'FORMATTING PERCENT CHANGE AND SETTING TOTALSTOCKVOLUME
  Sht.Cells(SummaryRow, 11) = Format(PercentChange, "0.00%")
  Sht.Cells(SummaryRow, 12) = TotalStockVolume
'CALCULATING ADDITIONAL SUMMARY ON GREATEST INCREASE DECREASE AND STOCKVOLUME
  If (GTPercentIncrease < PercentChange) Then
   GTPercentIncrease = PercentChange
   GTPercentInTicker = Sht.Cells(SummaryRow, 9)
  End If
  If (GTPercentDecrease > PercentChange) Then
   GTPercentDecrease = PercentChange
   GTPercentDeTicker = Sht.Cells(SummaryRow, 9)
  End If
  If (GTStockvolume < TotalStockVolume) Then
   GTStockvolume = TotalStockVolume
   GTStockVolTicker = Sht.Cells(SummaryRow, 9)
  End If
'RESETTING THE NEEDED VALUES FOR SUMMARY
  TotalStockVolume = 0
  SummaryRow = SummaryRow + 1
  StartValue = Sht.Cells(I + 1, 3).Value
 End If
Next I
'SETTING THE ADDITIONAL SUMMARY ON GREATEST INCREASE DECREASE AND STOCKVOLUME
  Sht.Cells(2, 15) = GTPercentInTicker
  Sht.Cells(2, 16) = Format(GTPercentIncrease, "0.00%")
  Sht.Cells(3, 15) = GTPercentDeTicker
  Sht.Cells(3, 16) = Format(GTPercentDecrease, "0.00%")
  Sht.Cells(4, 15) = GTStockVolTicker
  Sht.Cells(4, 16) = GTStockvolume
'RESETTING THE ADDITIONAL SUMMARY VARIABLES
GTPercentInTicker = ""
GTPercentDeTicker = ""
GTStockVolTicker = ""
GTPercentIncrease = 0
GTPercentDecrease = 0
GTStockvolume = 0
Next J
End Sub
