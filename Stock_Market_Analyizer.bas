Attribute VB_Name = "Module1"
Option Explicit
Sub Stock_Market_Workbook()

'When F5 is pressed this is the subroutine to start with.
'Main subroutine to loop through the worksheets

Dim wrksht As Worksheet

Application.ScreenUpdating = False

'Form created to inform user that program is running
WaitPls.Show vbModeless
WaitPls.Repaint

For Each wrksht In Worksheets
    wrksht.Select
    
    'Subroutine which analyizes each stock and creates the summary of each stock
    Call SM_Analyst
    
Next
Application.ScreenUpdating = True

'Return to first worksheet
Worksheets(1).Activate
Unload WaitPls
MsgBox ("Analysis Complete")
End Sub
Sub SM_Analyst()

'Subroutine which loops through dataset to considate data into unique tickers and creates summary

Dim Current_Ticker As String
Dim Total_Stk_Volume As Double
Dim Stock_Open_Price As Double
Dim Stock_close_price As Double

Dim Row_Num As Long
Dim Row_Num_Yearly As Long


Row_Num = 2
Row_Num_Yearly = 1

'Calling the Sort_Sheet subroutine

Call Sort_Sheet

'Do While loop is used instead of for loop to maxrows

Do While (Cells(Row_Num, 1).Value <> "")
    Current_Ticker = Cells(Row_Num, 1).Value
    Stock_Open_Price = Cells(Row_Num, 3).Value
    Total_Stk_Volume = Cells(Row_Num, 7).Value
    
'Do while loop summarizes like Tickers
    
    Do While (Current_Ticker = Cells(Row_Num, 1).Value)
        Row_Num = Row_Num + 1
        Total_Stk_Volume = Total_Stk_Volume + Cells(Row_Num, 7).Value
    Loop
    Stock_close_price = Cells(Row_Num - 1, 6).Value
    Row_Num_Yearly = Annual_Ticker_Summary(Row_Num_Yearly, Current_Ticker, Stock_Open_Price, Stock_close_price, Total_Stk_Volume)
Loop

'Subroutine which creates the Overall Stock Performance Table
Call Overall_Stock_Performance

End Sub
Sub Sort_Sheet()

' Subroutine to sort dataset into correct sort order even though this dataset is sorted already.
' Just in case subroutine()
'   Ticker Column ("A") is the key sort field
'       Date Column ("B") is the secondary sort field within column ("A")

    Columns("A:G").Sort key1:=Columns("A"), Order1:=xlAscending, key2:=Columns("B"), order2:=xlAscending, Header:=xlYes

End Sub

Function Annual_Ticker_Summary(Row_Num_Yearly As Long, Current_Ticker As String, Stock_Open_Price As Double, Stock_close_price As Double, Total_Stk_Volume As Double) As Long

'This function creates the second dataset for the Annual Ticker Summary.
'Returns row number which hold the place in the dataset

Dim Yearly_Change As Double
Dim Yearly_Change_Percent As Double

If (Row_Num_Yearly = 1) Then
    'Format Column Widths
    Columns("I").ColumnWidth = 7
    Columns("J").ColumnWidth = 12
    Columns("K").ColumnWidth = 16
    Columns("L").ColumnWidth = 17
    
    'Create Header
    Range("I1").Font.Bold = True
    Range("J1").Font.Bold = True
    Range("K1").Font.Bold = True
    Range("L1").Font.Bold = True
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
End If
Yearly_Change = Stock_close_price - Stock_Open_Price
If Yearly_Change > 0 Then
    Cells(Row_Num_Yearly + 1, 10).Interior.Color = vbGreen
    ElseIf Yearly_Change < 0 Then
        Cells(Row_Num_Yearly + 1, 10).Interior.Color = vbRed
End If
Cells(Row_Num_Yearly + 1, 9).Value = Current_Ticker
Cells(Row_Num_Yearly + 1, 10).Value = Yearly_Change
Cells(Row_Num_Yearly + 1, 11).NumberFormat = "0.00%"
If Stock_Open_Price = 0 Then
    Cells(Row_Num_Yearly + 1, 11).Value = 0
Else
    Yearly_Change_Percent = Yearly_Change / Stock_Open_Price
    Cells(Row_Num_Yearly + 1, 11).Value = Yearly_Change_Percent
End If
Cells(Row_Num_Yearly + 1, 12).NumberFormat = "###,###,###"
Cells(Row_Num_Yearly + 1, 12).Value = Total_Stk_Volume

'Annual_Ticker_Summary is the function return value for the row position

Annual_Ticker_Summary = Row_Num_Yearly + 1
End Function
Sub Overall_Stock_Performance()

'Subroutine which loops through the Annual Ticker Summary dataset to create the Overall Stock Performance

Dim GPI As Double
Dim GPI_Ticker As String
Dim GPD As Double
Dim GPD_Ticker As String
Dim GTV As Double
Dim GTV_Ticker As String

Dim Row_Num As Long

Row_Num = 2
GPI = Cells(Row_Num, 11).Value
GPD = Cells(Row_Num, 11).Value
GTV = Cells(Row_Num, 12).Value

' Decided to use a do while loop instead of a for loop to maxrows
' Loops through to determince Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total Volume

Do While (Cells(Row_Num, 9) <> "")
    If Cells(Row_Num, 11).Value > GPI Then
        GPI = Cells(Row_Num, 11).Value
        GPI_Ticker = Cells(Row_Num, 9).Value
    End If
    If Cells(Row_Num, 11) < GPD Then
        GPD = Cells(Row_Num, 11).Value
        GPD_Ticker = Cells(Row_Num, 9).Value
    End If
    If Cells(Row_Num, 12) > GTV Then
        GTV = Cells(Row_Num, 12).Value
        GTV_Ticker = Cells(Row_Num, 9).Value
    End If
    Row_Num = Row_Num + 1
Loop

'Creates Row and Column Headers for results

Range("N2").ColumnWidth = 25
Range("N2").Font.Bold = True
Range("N2").Value = "Greatest % Increase"

Range("N3").ColumnWidth = 25
Range("N3").Font.Bold = True
Range("N3").Value = "Greatest % Decrease"

Range("N4").ColumnWidth = 25
Range("N4").Font.Bold = True
Range("N4").Value = "Greatest Volume"

Range("O1").ColumnWidth = 7
Range("O1").Font.Bold = True
Range("O1").Value = "Ticker"

Range("P1").ColumnWidth = 17
Cells(1, 16).HorizontalAlignment = xlCenter
Range("P1").Font.Bold = True
Range("P1").Value = "Value"

Range("P2").Font.Color = vbGreen
Range("O2").Value = GPI_Ticker
Range("P2").NumberFormat = "0.00%"
Range("P2").Value = GPI

Range("P3").Font.Color = vbRed
Range("O3").Value = GPD_Ticker
Range("P3").NumberFormat = "0.00%"
Range("P3").Value = GPD

Range("O4").Value = GTV_Ticker
Range("P4").NumberFormat = "###,###,###"
Range("P4").Value = GTV

End Sub
