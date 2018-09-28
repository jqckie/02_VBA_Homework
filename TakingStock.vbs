Attribute VB_Name = "Module1"
Sub TakingStock()

'***************************************************
'* ____                                            *
'*  |    _   |   .   _    _   (   |_   _    _  |   *
'*  |   (_\  |<  |  | |  (_|  _)  |_  (_)  (_  |<  *
'*                        _/                       *
'***************************************************
'Upon running once, this script will loop through  '
'each year of stock data on every worksheet in the '
'active workbook and grab the total amount of      '
'volume each stock had over the year, display the  '
'ticker symbol to coincide with the total volume,  '
'capture yearly change from what the stock opened  '
'the year at to what the closing price was, capture'
'the percent change from the what it opened the    '
'year at to what it closed, locate the stocks with '
'the "Greatest % increase/Decrease/total volume,"  '
'and also features conditional formatting that will'
'highlight positive yearly changes in green and    '
'negative yearly changes in red.                   '
'===================================================
'               CREATE VARIABLES                   '
'===================================================
Dim Tick As String
Dim GIncTick As String
Dim GDecTick As String
Dim GTSVTick As String
Dim Volume As Double
Dim TSV As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChg As Double
Dim GInc As Double
Dim GDec As Double
Dim GTSV As Double
Dim WS_Count As Integer
Dim LastRow As Double
Dim i As Integer
Dim j As Double
Dim k As Integer
Dim TableRow As Integer
Dim CurrentProgress As Double
Dim ProgressPercentage As Double
Dim BarWidth As Long

'===================================================
'               PREPARE FOR THE WORK               '
'===================================================
'Display macro status message
Application.StatusBar = "SUMMARIZING DATA"

'Calculate precise values
ActiveWorkbook.PrecisionAsDisplayed = False

'Count sheets in active workbook
WS_Count = ActiveWorkbook.Worksheets.Count
    
'Activate the Macro Progress status bar
Call InitProgressBar

'===================================================
'  BEGIN THE LOOP FOR EACH SHEET IN THIS WORKBOOK  '
'===================================================
For i = 1 To WS_Count
    Worksheets(i).Activate
    With ActiveSheet
    'Count Number of Rows on Active Sheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
            
    'Set TableRow to start on 2nd row
    TableRow = 1

'===================================================
' BEGIN THE LOOP FOR EACH ACTIVE ROW ON THIS SHEET '
'===================================================
    For j = 2 To (LastRow + 1)
                    
'>>>>>>>>>>>>>>>'First row for this ticker?
                If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
                    
                    'Set Ticker final values from previous row
                    If TableRow >= 2 Then
                        
                        'Set Close
                        ClosePrice = Cells(j - 1, 6).Value
                        
                        'Set Yearly Change & print to table
                        YearlyChange = ClosePrice - OpenPrice
                        Cells(TableRow, 10).Value = YearlyChange
                                                                                             
                        'Set Percent Change & print to table
                        If OpenPrice = 0 Then
                            PercentChg = 0
                            Else
                            PercentChg = (YearlyChange / OpenPrice)
                        End If
                        Cells(TableRow, 11).Value = PercentChg
                        Cells(TableRow, 11).NumberFormat = "0.00%"
                        
                        'Print Total Stock Volume to table
                        Cells(TableRow, 12).Value = TSV
                        
                        'Check for Greatest values
                            If PercentChg > GInc Then
                                GInc = PercentChg
                                GIncTick = Tick
                            End If
                            If PercentChg < GDec Then
                                GDec = PercentChg
                                GDecTick = Tick
                            End If
                            If TSV > GTSV Then
                                GTSV = TSV
                                GTSVTick = Tick
                            End If
                    End If
                    
                    'Set Ticker values from current row
                    Tick = Cells(j, 1).Value            'Set ticker
                    OpenPrice = Cells(j, 3).Value       'Set Open
                    Volume = Cells(j, 7).Value          'Set Vol
                    TSV = Volume                        'Start running total for Total Stock Volume
                    TableRow = TableRow + 1             'Start new row on table
                    Cells(TableRow, 9).Value = Tick     'Populate new row on table
                    
'>>>>>>>>>>>>>>>'Continuation for this ticker?
                ElseIf Cells(j, 1).Value = Cells(j - 1, 1).Value Then
                    Volume = Cells(j, 7).Value
                    TSV = TSV + Volume
                End If

'===================================================
'               UPDATE PROGRESS BAR                '
'===================================================
    CurrentProgress = ((i - 1) / WS_Count) + ((j / LastRow) * (1 / WS_Count))
    BarWidth = Progress.Border.Width * CurrentProgress
    ProgressPercentage = Round(CurrentProgress * 100, 0)
    
    Progress.Bar.Width = BarWidth
    Progress.Text.Caption = ProgressPercentage & "% Complete"
    
    DoEvents
    
    Next j                                 'NEXT ROW
            
'===================================================
'               CREATE SUMMARY TABLES              '
'===================================================
Range("I1").Value = "Ticker"                            'Ticker
Range("J1").Value = "Yearly Change"                     'Yearly Change
Range("K1").Value = "Percent Change"                    'Percent Change
Range("L1").Value = "Total Stock Volume"                'Total Stock Volume
Range("O2").Value = "Greatest % Increase"               'Greatest % Increase
Range("O3").Value = "Greatest % Decrease"               'Greatest % Decrease
Range("O4").Value = "Greatest Total Volume"             'Greatest Total Volume
Range("P1").Value = "Ticker"                            'Ticker
Range("Q1").Value = "Value"                             'Value
    
'===================================================
'         POPULATE GREATEST SUMMARY TABLE          '
'===================================================
'Greatest % Increase
Range("P2").Value = GIncTick
Range("Q2").Value = GInc
Range("Q2").NumberFormat = "0.00%"
'Greatest % Decrease
Range("P3").Value = GDecTick
Range("Q3").Value = GDec
Range("Q3").NumberFormat = "0.00%"
'Greatest Total Volume
Range("P4").Value = GTSVTick
Range("Q4").Value = GTSV
    
'===================================================
'               AUTOFIT TABLE COLUMNS              '
'===================================================
Columns("J:L").EntireColumn.AutoFit
Columns("O:Q").EntireColumn.AutoFit
    
'===================================================
'               FORMAT YEARLY CHANGE               '
'===================================================
    For k = 2 To (TableRow - 1)
    Cells(k, 10).NumberFormat = "0.000000000"
        If Cells(k, 10).Value >= 0 Then
            Cells(k, 10).Interior.ColorIndex = 4 'Fill Green
            Else: Cells(k, 10).Interior.ColorIndex = 3 'Fill Red
        End If
    Next k

'===================================================
'               RESET FOR NEXT SHEET               '
'===================================================
'Clear Greatest values from current sheet
    GInc = 0
    GDec = 0
    GTSV = 0
'Leave each tab on A1
ActiveSheet.Range("A1").Select

'Show progress for number of sheets processed
Application.StatusBar = i & " OUT OF " & WS_Count & " SHEETS SUMMARIZED"
             
'***********CURRENT SHEET LOOP COMPLETE***********

Next i                                   'NEXT SHEET
    
'===================================================
'                 COMPLETE THE WORK                '
'===================================================
'Display macro status message
Application.StatusBar = "SUMMARY COMPLETE"

'Leave workbook open to first tab
Worksheets(1).Activate
    
'Stop displaying Macro Progress bar
Unload Progress

'Display macro status message
Application.StatusBar = "Ready"

End Sub

Sub InitProgressBar()

With Progress
    
    .Bar.Width = 0
    .Text.Caption = "0% Complete"
    .Show vbModeless
    

End With

End Sub
