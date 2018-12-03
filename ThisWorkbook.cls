VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock()
    
    'Loop through each existing sheet except "UniqueTicker"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "UniqueTicker" Then
            Call ProcessData(ws.Name)
       End If
    Next ws

    'After processing is done, we delete the "UniqueTicker" sheet
    Application.DisplayAlerts = False
    Worksheets("UniqueTicker").Delete
    Application.DisplayAlerts = True
    
End Sub


Sub ProcessData(wsName As String)
  
    Dim uniquesTickerArray As Variant
    Dim ticker As String, sumVol As Double
    Dim FirstOpeningPrice, LastClosingPrice, YearlyChange As Double
    
    With Worksheets(wsName)
        ' --------------------------------------------
        'Set Header
        ' --------------------------------------------
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
        
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        
        
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
    
        .Range("Q2").NumberFormat = "0.00%"
        .Range("Q3").NumberFormat = "0.00%"
        .Range("Q4").NumberFormat = "0"
    End With
    
    ' ------------------------------------------------------
    ' This funtion returns an array of all the unique Ticker
    ' ------------------------------------------------------
    uniquesTickerArray = Filter_UniquesTicker(wsName)

    For i = 1 To UBound(uniquesTickerArray)

        ticker = uniquesTickerArray(i, 1)

        With Worksheets(wsName)
            ' --------------------------------------------
            ' Filter Data by Ticker
            ' --------------------------------------------
            .Range("A:A").AutoFilter Field:=1, Criteria1:=ticker

             Set volumeRange = .Range("G2", .Range("G2").End(xlDown)).SpecialCells(xlCellTypeVisible)
             Set openPriceRange = .Range("C2", .Range("C2").End(xlDown)).SpecialCells(xlCellTypeVisible)
             
             lRow = .Cells(Rows.Count, 1).End(xlUp).Row
             FirstOpeningPrice = openPriceRange(1, 1)
             LastClosingPrice = .Range("F" & lRow).Value
             YearlyChange = LastClosingPrice - FirstOpeningPrice

             'Ticker
             .Range("I" & i + 1) = ticker
             'Yearly Change
             .Range("J" & i + 1).NumberFormat = "0.000000000"
             .Range("J" & i + 1) = YearlyChange
             .Range("J" & i + 1).Interior.ColorIndex = IIf(YearlyChange < 0, 3, 4) '3 =Red, 4=Green
             'Percent Change
             If FirstOpeningPrice > 0 Then
               .Range("K" & i + 1) = YearlyChange / FirstOpeningPrice
             Else
              .Range("K" & i + 1) = YearlyChange
             End If
             
             .Range("K" & i + 1).NumberFormat = "0.00%"
             'Total Stock Volume
             .Range("L" & i + 1) = WorksheetFunction.Sum(volumeRange)
             .Range("L" & i + 1).NumberFormat = "0"
        End With

    Next

    Worksheets(wsName).ShowAllData
    
    Dim gI, gD, maxVolume As Double

    With Worksheets(wsName)
        'Greatest % Increase
        gI = Application.WorksheetFunction.Max(.Range("K:K"))
        gI = Format(gI, "0.00%")
        Set gCell = .Range("K:K").Find(what:=gI, LookIn:=xlValues, lookat:=xlWhole)

        .Range("P2").Value = .Range("I" & gCell.Row).Value
        .Range("Q2").Value = gI

        'Greatest % Decrease
        gD = Application.WorksheetFunction.Min(.Range("K:K"))
        gD = Format(gD, "0.00%")
        Set dCell = .Range("K:K").Find(what:=gD, LookIn:=xlValues, lookat:=xlWhole)

        .Range("P3").Value = .Range("I" & dCell.Row).Value
        .Range("Q3").Value = gD

        'Greatest Total Volume
        maxVolume = Application.WorksheetFunction.Max(.Range("L:L"))
        Set maxCell = .Range("L:L").Find(what:=maxVolume, LookIn:=xlValues)

        .Range("P4").Value = .Range("I" & maxCell.Row).Value
        .Range("Q4").Value = maxVolume
    End With

End Sub

' ------------------------------------------------------
' This funtion returns an array of unique Tickers
' ------------------------------------------------------
Function Filter_UniquesTicker(wsName As String) As Variant
    
    Dim uniqueSheet As String
    uniqueSheet = "UniqueTicker"
    Dim wsUnique As Worksheet
    
    '"UniqueTicker" sheet is an extra place holder to copy the unique ticker data over.
    'If "UniqueTicker" sheet does not exsist then we create it.
    If Not WorksheetExists(uniqueSheet) Then
        Set wsUnique = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsUnique.Name = uniqueSheet
    Else
        Set wsUnique = Worksheets(uniqueSheet)
    End If
    
  
    Dim uniquesArray As Variant
    Dim lastRow As Long
    With wsUnique
        Worksheets(wsName).Columns("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("A1"), Unique:=True
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        'Save "UniqueTicker" data into an array
        uniquesArray = .Range("A2:A" & lastRow)
        .Columns("A").ClearContents
    End With
    
    Filter_UniquesTicker = uniquesArray
    
End Function

' ---------------------------------------------------------------
' This funtion checkes to see a sheet with the name exists or not
' ---------------------------------------------------------------
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
     Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     WorksheetExists = Not sht Is Nothing
End Function










