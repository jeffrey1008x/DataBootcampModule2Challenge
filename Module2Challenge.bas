Attribute VB_Name = "Module11"
Sub assignment2()

    Dim ws As Worksheet
    Dim last_row As Long
    Dim last_col As Long
    Dim ticker As String
    Dim line_no As Long
    Dim max_vol As LongLong
    
    Dim lookup As Object
    'Set lookup = CreateObject("Scripting.Dictionary")

    ' Loop through each sheet
    For Each ws In ThisWorkbook.Worksheets
    
        Set lookup = CreateObject("Scripting.Dictionary")

        last_row = ws.Range("A1").End(xlDown).Row
        last_col = ws.Range("A1").End(xlToRight).Column
        
        ' Store important information for each ticker
        For i = 2 To last_row
            ticker = ws.Cells(i, 1).Value
            If lookup.Exists(ticker) = True Then
                lookup(ticker)("total") = lookup(ticker)("total") + ws.Cells(i, 7).Value
                
                If ws.Cells(i, 2).Value < lookup(ticker)("start") Then
                    lookup(ticker)("start") = ws.Cells(i, 2).Value
                    lookup(ticker)("start_price") = ws.Cells(i, 3).Value
                End If
                
                If ws.Cells(i, 2).Value > lookup(ticker)("end") Then
                    lookup(ticker)("end") = ws.Cells(i, 2).Value
                    lookup(ticker)("end_price") = ws.Cells(i, 6).Value
                End If
            Else:
                lookup.Add ticker, CreateObject("Scripting.Dictionary")
                lookup(ticker).Add "total", ws.Cells(i, 7).Value
                
                lookup(ticker).Add "start", ws.Cells(i, 2).Value
                lookup(ticker).Add "end", ws.Cells(i, 2).Value

                lookup(ticker).Add "start_price", ws.Cells(i, 3).Value
                lookup(ticker).Add "end_price", ws.Cells(i, 6).Value
                
                
            End If
        Next i
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        line_no = 2
        max_vol = 0
        
        ' Give the max values an initial value
        ws.Cells(2, 16).Value = lookup.keys()(0)
        ws.Cells(2, 17).Value = FormatPercent(ws.Range("K2").Value, 2)
        ws.Cells(3, 16).Value = lookup.keys()(0)
        ws.Cells(3, 17).Value = FormatPercent(ws.Range("K2").Value, 2)
        
        ' Calculate the changes for each ticker
        For Each k In lookup.keys

            ws.Cells(line_no, 9).Value = k
            ws.Cells(line_no, 10).Value = lookup(k)("end_price") - lookup(k)("start_price")
            If ws.Cells(line_no, 10).Value > 0 Then
                ws.Cells(line_no, 10).Interior.Color = RGB(0, 255, 0)
            ElseIf ws.Cells(line_no, 10).Value < 0 Then
                ws.Cells(line_no, 10).Interior.Color = RGB(255, 0, 0)
            End If
            ws.Cells(line_no, 11).Value = FormatPercent(ws.Cells(line_no, 10).Value / lookup(k)("start_price"), 2)
            
            If ws.Cells(line_no, 11).Value > 0 Then
                ws.Cells(line_no, 11).Interior.Color = RGB(0, 255, 0)
            ElseIf ws.Cells(line_no, 11).Value < 0 Then
                ws.Cells(line_no, 11).Interior.Color = RGB(255, 0, 0)
            End If
            
            ws.Cells(line_no, 12).Value = lookup(k)("total")
            
            
            If lookup(k)("total") > max_vol Then
                max_vol = lookup(k)("total")
                ws.Cells(4, 16).Value = k
                ws.Cells(4, 17).Value = max_vol
            End If
            
            line_no = line_no + 1
            
            
        Next k
        
        ' Now find the Greatest Increase and Decrease
        For j = 2 To ws.Range("K2").End(xlDown).Row
            If ws.Cells(j, 11).Value > ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(2, 17).Value = FormatPercent(ws.Cells(j, 11).Value, 2)
            End If
        Next j
        
        For j = 2 To ws.Range("K2").End(xlDown).Row
            If ws.Cells(j, 11).Value < ws.Cells(3, 17).Value Then
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                ws.Cells(3, 17).Value = FormatPercent(ws.Cells(j, 11).Value, 2)
            End If
        Next j

    Next ws

End Sub
