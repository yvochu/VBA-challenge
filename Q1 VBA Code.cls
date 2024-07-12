VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub QuarterlyCalc()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim data As Variant
    Dim i As Long
    Dim ticker As String
    Dim tickerStartRow As Long
    Dim quarterStart As Date
    Dim quarterEnd As Date
    Dim openingPrice As Double

    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    outputRow = 2

    Set ws = ThisWorkbook.Sheets("Q1")
    
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    

   
    
    
    data = ws.Range("A1:G" & lastRow).Value 'putting my data in an array
    For i = 2 To lastRow 'Second row to last row
        
       
        If data(i, 1) <> ticker Then
            'quarterly change
            closingPrice = data(i - 1, 6)
            If openingPrice <> 0 Then
                quarterlyChange = closingPrice - openingPrice
                percentageChange = (closingPrice - openingPrice) / openingPrice
            Else
                quarterlyChange = 0
                percentageChange = 0
            End If
            
            
            ws.Cells(outputRow, 14).Value = ticker ' Ticker name in column N
            ws.Cells(outputRow, 15).Value = quarterlyChange ' Quarterly change in column O
            ws.Cells(outputRow, 16).Value = percentageChange ' Percentage change in column P
            ws.Cells(outputRow, 17).Value = totalVolume ' Total volume in column Q
            
            
            outputRow = outputRow + 1
            
            
            ticker = data(i, 1)
            tickerStartRow = i
            quarterStart = DateSerial(Year(data(i, 2)), ((Month(data(i, 2)) - 1) \ 3) * 3 + 1, 1)
            quarterEnd = DateAdd("m", 3, quarterStart) - 1
            openingPrice = data(i, 3)
            totalVolume = 0
        End If
        
        
        totalVolume = totalVolume + data(i, 7)
        
        
        If data(i, 2) > quarterEnd Then
            'quarterly change
            closingPrice = data(i - 1, 6)
            If openingPrice <> 0 Then
                quarterlyChange = closingPrice - openingPrice
                percentageChange = (closingPrice - openingPrice) / openingPrice
            Else
                quarterlyChange = 0
                percentageChange = 0
            End If
            
           
            ws.Cells(outputRow, 14).Value = ticker ' Ticker name in column N
            ws.Cells(outputRow, 15).Value = quarterlyChange ' Quarterly change in column O
            ws.Cells(outputRow, 16).Value = percentageChange ' Percentage change in column P
            ws.Cells(outputRow, 17).Value = totalVolume ' Total volume in column Q
            
            
            outputRow = outputRow + 1
            
          
            tickerStartRow = i
            quarterStart = DateSerial(Year(data(i, 2)), ((Month(data(i, 2)) - 1) \ 3) * 3 + 1, 1)
            quarterEnd = DateAdd("m", 3, quarterStart) - 1
            openingPrice = data(i, 3)
            totalVolume = 0
        End If
        
    Next i
    
    
    closingPrice = data(lastRow, 6)
    If openingPrice <> 0 Then
        quarterlyChange = closingPrice - openingPrice
        percentageChange = (closingPrice - openingPrice) / openingPrice
    Else
        quarterlyChange = 0
        percentageChange = 0
    End If
    ws.Cells(outputRow, 14).Value = ticker ' Ticker name in column N
    ws.Cells(outputRow, 15).Value = quarterlyChange ' Quarterly change in column O
    ws.Cells(outputRow, 16).Value = percentageChange ' Percentage change in column P
    ws.Cells(outputRow, 17).Value = totalVolume ' Total volume in column Q
    
   
    ws.Range("N1:Q1").Value = Array("Ticker", "Quarterly Change", "Percent Change", "Total Volume")
    
    
   

End Sub

