Attribute VB_Name = "Module1"
Sub VBA_Project():
'Define our variables
Dim ticker As String
Dim opening As Double
Dim high As Double
Dim Low As Double
Dim closing As Double
Dim Volume As LongLong
Dim LR As Long
Dim Yearly_change As Double
Dim Percentage_change As Double
Dim WS_Count As Integer
Dim ws As Worksheet
Dim Ticker_header As String
Dim LR_Count As Long
Dim Max As Double
Dim Min As Double
Dim MaxV As LongLong
Dim Tickers As String
Dim Col As New Collection
Dim ValCell As String
Dim i As LongLong
Dim n As Integer
Dim lastrow_Ticker As LongLong
Dim lastrow_A As Long
Dim rowfirst  As Long
Dim rowlast As LongLong


Dim unique As String

'for loop


For Each ws In Worksheets
LR_Ticker = ws.Cells(Rows.Count, 9).End(xlUp).Row
LR_Count = LR_Ticker

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    ws.Range("A2:A" & lastrow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("I2"), unique:=True
    
For i = 2 To LR_Ticker

If ws.Cells(i, 9) <> "" Then

rowfirst = ws.Columns(1).Find(What:=ws.Cells(i, 9).Value, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False).Row


rowlast = ws.Columns(1).Find(What:=ws.Cells(i, 9).Value, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False).Row


opening = ws.Cells(rowfirst, 3)


closing = ws.Cells(rowlast, 6)


Yearly_change = closing - opening
ws.Cells(i, 10).Value = Yearly_change

Percentage_chage = Yearly_change / opening
ws.Cells(i, 11).Value = Percentage_chage

ws.Range("K" & i).NumberFormat = "0.00%"

Volume = WorksheetFunction.Sum(ws.Range((ws.Cells(rowfirst, 7)), (ws.Cells(rowlast, 7))))
ws.Cells(i, 12).Value = Volume
ws.Range("L" & i).NumberFormat = "0"


'triming up some extra calculations on blanks

End If
If ws.Cells(i, 9) = "" Then
    ws.Cells(i, 10) = ""
    ws.Cells(i, 11) = ""
    ws.Cells(i, 12) = ""
    
End If




' formatting color with if
If ws.Cells(i, 10).Value > 0 Then

ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
   
End If

'finding the max value and putting it in its place
Max = WorksheetFunction.Max(ws.Range("K:K"))

ws.Cells(2, 17).Value = Max
ws.Cells(2, 17).NumberFormat = "0.00%"

'finding the min value and putting in its place
Min = WorksheetFunction.Min(ws.Range("K:K"))

ws.Cells(3, 17).Value = Min
ws.Cells(3, 17).NumberFormat = "0.00%"

'Finding the max volumne and put in its place
MaxV = WorksheetFunction.Max(ws.Range("L:L"))

ws.Cells(4, 17).Value = MaxV
ws.Cells(4, 17).NumberFormat = "0"

If ws.Cells(i, 11).Value = Max Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
End If

If ws.Cells(i, 11).Value = Min Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
End If

If ws.Cells(i, 12).Value = MaxV Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
End If


Next i

'headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"




Next ws







End Sub
