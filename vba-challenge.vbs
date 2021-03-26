Sub Stockdata()
    Dim ws As Worksheet
   
    For Each ws In Worksheets
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
                
        Dim ticker As String
        ticker = " "
        Dim totalticker As Double
        totalticker = 0
        Dim openp As Double
        openp = 0
        Dim closep As Double
        closep = 0
        Dim yearlychange As Double
        yearlychange = 0
        Dim yearlychangepercent As Double
        yearlychangepercent = 0
        
        Dim maxticker As String
        maxticker = " "
        Dim minticker As String
        minticker = " "
        Dim maxpercent As Double
        maxpercent = 0
        Dim minpercent As Double
        minpercent = 0
        Dim maxvolumeticker As String
        maxvolumeticker = " "
        Dim maxvolume As Double
        maxvolume = 0
        
        Dim summarytable As Long
        summarytable = 2
        
        Dim Lastrow As Long
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        openp = ws.Cells(2, 3).Value
                       
        For i = 2 To Lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closep = ws.Cells(i, 6).Value
                yearlychange = closep - openp
                
                If openp <> 0 Then
                yearlychangepercent = (yearlychange / openp) * 100
                End If
            
            totalticker = totalticker + ws.Cells(i, 7).Value
            ws.Range("I" & summarytable).Value = ticker
            ws.Range("J" & summarytable).Value = yearlychange
            
                If (yearlychange > 0) Then
                ws.Range("J" & summarytable).Interior.ColorIndex = 4
            
                ElseIf (yearlychange <= 0) Then
                ws.Range("J" & summarytable).Interior.ColorIndex = 3
                End If
        
            ws.Range("K" & summarytable).Value = (CStr(yearlychangepercent) & "%")
            ws.Range("L" & summarytable).Value = totalticker
        
            summarytable = summarytable + 1
            yearlychange = 0
            closep = 0
            openp = ws.Cells(i + 1, 3).Value
        
                If (yearlychangepercent > maxpercent) Then
                maxpercent = yearlychangepercent
                maxticker = ticker
                ElseIf (yearlychangepercent < minpercent) Then
                minpercent = yearlychangepercent
                minticker = ticker
                End If
        
                If (totalticker > maxvolume) Then
                maxvolume = totalticker
                maxvolumeticker = ticker
                End If
        
                yearlychangepercent = 0
                totalticker = 0
        
            Else
            totalticker = totalticker + ws.Cells(i, 7).Value
            End If
    
                
                
            
        Next i
        
 
    
        ws.Range("Q2").Value = (CStr(maxpercent) & "%")
        ws.Range("Q3").Value = (CStr(minpercent) & "%")
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("P2").Value = maxticker
        ws.Range("P3").Value = minticker
        ws.Range("Q4").Value = maxvolume
        ws.Range("P4").Value = mavolumeticker
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
Next ws


End Sub
