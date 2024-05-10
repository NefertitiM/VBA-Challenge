Attribute VB_Name = "Module1"
Sub stockData()

    For Each ws In Worksheets
    
    
        Dim i As Long
    
        Dim ticker As String
    
        Dim Row_Number As Integer
        Row_Number = 2
    
        Dim total As Double
        total = 0
    
        Dim openNum As Double
        openNum = 0
    
        Dim closeNum As Double
        closeNum = 0
    
        
    
        Dim WorksheetName As String
        
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        WorksheetName = ws.Name
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
    
    
    
            For i = 2 To lastRow
    
        
                ' Last row of the same ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                
                ' Set variables
                ticker = ws.Cells(i, 1).Value
                total = total + ws.Cells(i, 7).Value
                closeNum = closeNum + ws.Cells(i, 6).Value
                
        
                'Add to summary table
                ws.Range("I" & Row_Number).Value = ticker
                ws.Range("L" & Row_Number).Value = total
                ws.Range("J" & Row_Number).Value = closeNum - openNum
                ws.Range("K" & Row_Number).Value = ((closeNum / openNum) - 1)
                ws.Range("K" & Row_Number).NumberFormat = "0.00%"
                
                
               
        
                'Reset Variables
                Row_Number = Row_Number + 1
                total = 0
                openNum = 0
                closeNum = 0
                
                'If first value is 0
                ElseIf openNum = 0 Then
                
                openNum = openNum + ws.Cells(i, 3).Value
                total = total + ws.Cells(i, 7).Value
                closeNum = closeNum + 0
                
                
                
                
                'In same ticker
                Else
                total = total + ws.Cells(i, 7).Value
                closeNum = closeNum + 0
                
        
                End If
    
    
    Next i


    'PartII
    
    Dim lastRow2 As Long
    lastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
        For p = 2 To lastRow2
            
            If ws.Cells(p, 10).Value < 0 Then
                ws.Cells(p, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(p, 10).Interior.ColorIndex = 4
            End If
        Next p
    
    
    
    
    
    
    
    
    'PartIII
    
    
    
    Dim maxPercent As Double
    Dim minPercent As Double
    Dim maxVolume As Double
    
    Dim maxPercentticker As String
    Dim minPercentticker As String
    Dim maxVolumeticker As String
    
    maxPercent = WorksheetFunction.Max(ws.Range("K2:K4"))
    minPercent = WorksheetFunction.Min(ws.Range("K2:K4"))
    maxVolume = WorksheetFunction.Max(ws.Range("L2:L4"))
    
    
    For j = 2 To lastRow2
    
        If ws.Cells(j, 11).Value = maxPercent Then
        
            ws.Range("Q2").Value = maxPercent
            ws.Range("Q2").NumberFormat = "0.00%"
            maxPercentticker = ws.Cells(j, 9).Value
            ws.Range("P2").Value = maxPercentticker
            
            maxPercentticker = ""
        
        ElseIf ws.Cells(j, 11).Value = minPercent Then
            
            ws.Range("Q3").Value = minPercent
            ws.Range("Q3").NumberFormat = "0.00%"
            minPercentticker = ws.Cells(j, 9).Value
            ws.Range("P3").Value = minPercentticker
            
            minPercentticker = ""
            
        
        Else
            
        End If
    
    Next j
    
    
    
    
    
    For k = 2 To lastRow2
        
        If ws.Cells(k, 12).Value = maxVolume Then
            
            ws.Range("Q4").Value = maxVolume
            ws.Range("Q4").NumberFormat = "0.00E+00"
            maxVolumeticker = ws.Cells(k, 9).Value
            ws.Range("P4").Value = maxVolumeticker
            maxVolumeticker = ""
            
        End If
        
    Next k
            
            
                
    Next ws


End Sub

