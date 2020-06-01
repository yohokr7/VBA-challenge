Attribute VB_Name = "Module1"
Sub RunWorksheets():
    Dim Sheet As Worksheet
    Application.ScreenUpdating = False
    For Each Sheet In Worksheets
        Sheet.Select
        Call Solution
    Next
    Application.ScreenUpdating = True

End Sub
Sub Solution():
    'Declaring Variables
    Dim RowIndex As Long
    Dim LastRow As Long
    
    Dim TickerSymbol As String
    Dim BeginOpen As Double
    Dim EndOpen As Double
    Dim YrOpenChng As Double
    Dim PerChange As Double
    Dim TotalVol As Double
    
    Dim SummaryTableRow As Integer
    Dim SummLastRow As Integer
    Dim HeaderRow As Integer
    Dim StartRow As Integer
    Dim IndexColumn As Integer
    Dim VolColumn As Integer
    Dim GreatColumn As Integer
    Dim TickerColumn As Integer
    Dim ValueColumn As Integer
    Dim GreatIncRow As Integer
    Dim GreatDecRow As Integer
    Dim GreatVolRow As Integer
    
    Dim PerRange As Range
    Dim VolRange As Range
    Dim GreatInc As Double
    Dim GreatDec As Double
    Dim GreatVol As Double
    Dim IncRow As Integer
    Dim DecRow As Integer
    Dim VolRow As Integer
    
    'Defining Variables
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TotalVol = 0
    
    HeaderRow = 1
    StartRow = 2
    SummaryTableRow = 2
    GreatIncRow = 2
    GreatDecRow = 3
    GreatVolRow = 4
    
    IndexColumn = 9
    YrChngColumn = 10
    PerChngColumn = 11
    VolColumn = 12
    GreatColumn = 15
    TickerColumn = 16
    ValueColumn = 17
    
    'Creating Headers
    Cells(HeaderRow, IndexColumn).Value = "Index"
    Cells(HeaderRow, YrChngColumn).Value = "Yearly Change"
    Cells(HeaderRow, PerChngColumn).Value = "Percent Change"
    Cells(HeaderRow, VolColumn).Value = "Total Stock Volume"
    Cells(HeaderRow, TickerColumn).Value = "Ticker"
    Cells(HeaderRow, ValueColumn).Value = "Value"
    Cells(GreatIncRow, GreatColumn).Value = "Greatest % Increase"
    Cells(GreatDecRow, GreatColumn).Value = "Greatest % Decrease"
    Cells(GreatVolRow, GreatColumn).Value = "Greatest Total Volume"
    
    'Fitting Columns
    Columns(YrChngColumn).AutoFit
    Columns(PerChngColumn).AutoFit
    Columns(GreatColumn).AutoFit
    
    
    'Running Analysis
    For RowIndex = StartRow To LastRow
    
            If Cells(RowIndex + 1, 1).Value <> Cells(RowIndex, 1).Value Then
                TickerSymbol = Cells(RowIndex, 1).Value
                TotalVol = TotalVol + Cells(RowIndex, 7).Value
                EndOpen = Cells(RowIndex, 3).Value
                YrOpenChng = EndOpen - BeginOpen
                
                If BeginOpen = 0 Then
                PerChange = 0
                Else
                PerChange = (EndOpen / BeginOpen) - 1
                End If
                
                Cells(SummaryTableRow, IndexColumn).Value = TickerSymbol
                Cells(SummaryTableRow, YrChngColumn).Value = YrOpenChng
                Cells(SummaryTableRow, PerChngColumn).Value = FormatPercent(PerChange, 2)
                Cells(SummaryTableRow, VolColumn).Value = TotalVol
                
                'Formating Cells
                If YrOpenChng >= 0 Then
                    Cells(SummaryTableRow, YrChngColumn).Interior.ColorIndex = 4
                Else
                    Cells(SummaryTableRow, YrChngColumn).Interior.ColorIndex = 3
                End If
                
                TotalVol = 0
                SummaryTableRow = SummaryTableRow + 1
                
            ElseIf Cells(RowIndex, 1).Value <> Cells(RowIndex - 1, 1).Value Then
                'Creating Beginning Opening Value
                BeginOpen = Cells(RowIndex, 3)
                TotalVol = TotalVol + Cells(RowIndex, 7).Value
            
            Else
                TotalVol = TotalVol + Cells(RowIndex, 7).Value
            End If
    Next RowIndex
    
    
    
    'Getting the Greatest
    SummLastRow = SummaryTableRow - 1
    
    Set PerRange = Range(Cells(StartRow, PerChngColumn), Cells(SummLastRow, PerChngColumn))
    Set VolRange = Range(Cells(StartRow, VolColumn), Cells(SummLastRow, VolColumn))
    GreatInc = Application.WorksheetFunction.Max(PerRange)
    GreatDec = Application.WorksheetFunction.Min(PerRange)
    GreatVol = Application.WorksheetFunction.Max(VolRange)
    
    'Get Row for Ticker
    IncRow = Application.WorksheetFunction.Match(GreatInc, PerRange, 0) + 1
    DecRow = Application.WorksheetFunction.Match(GreatDec, PerRange, 0) + 1
    VolRow = Application.WorksheetFunction.Match(GreatVol, VolRange, 0) + 1
    
    'Inputting Values
    Cells(GreatIncRow, ValueColumn).Value = FormatPercent(GreatInc)
    Cells(GreatIncRow, TickerColumn).Value = Cells(IncRow, IndexColumn).Value
    Cells(GreatDecRow, ValueColumn).Value = FormatPercent(GreatDec)
    Cells(GreatDecRow, TickerColumn).Value = Cells(DecRow, IndexColumn).Value
    Cells(GreatVolRow, ValueColumn).Value = GreatVol
    Cells(GreatVolRow, TickerColumn).Value = Cells(VolRow, IndexColumn).Value
    Columns(ValueColumn).AutoFit

End Sub
