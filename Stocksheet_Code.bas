Attribute VB_Name = "Module1"
Sub Stocksheet()

' Define the Current Ws as Worksheet
Dim ws As Worksheet
    
Dim SummaryHeader As Boolean

' Boolean for Defining the headers needed for the question, if summary headers are there(True) they do not need to be put in
SummaryHeader = False
        
    
' Loop through each worksheet in the workbook
For Each ws In Worksheets
    
        ' Define Tickername and empty string to be filled
        Dim TickerName As String
        TickerName = " "
        
        ' Set an initial variable for holding the total per ticker name
        Dim TickerTotal As Double
        TickerTotal = 0
        
        ' Set variables
        Dim OpenPrice As Double
        OpenPrice = 0
        Dim ClosePrice As Double
        ClosePrice = 0
        Dim PriceChange As Double
        PriceChange = 0
        Dim PercentChange As Double
        PercentChange = 0
    
         
        ' Keep track of  each ticker name in current ws
        Dim SummaryRow As Long
        SummaryRow = 2
        
        ' Allows me to go up to the last row
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Set Titles for Column Names
        If SummaryHeader Then
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
        Else
            'If the page has the 4 column titles change to True
            SummaryHeader = True
        End If
        
        ' Set initial value of Open Price for the first Ticker of ws
        OpenPrice = ws.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet(Row2) till its last row
        For i = 2 To Lastrow
        
      
            ' Check if we are still within the same ticker name, if not prints results to summary table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set the ticker name
                TickerName = ws.Cells(i, 1).Value
                
                ' Calculate close price and price drop/increase
                ClosePrice = ws.Cells(i, 6).Value
                PriceChange = ClosePrice - OpenPrice
                
                ' Need This to not get error
                If OpenPrice <> 0 Then
                    PercentChange = (PriceChange / OpenPrice) * 100
                Else
               
                End If
                
                ' Add to the Ticker name total count
                TickerTotal = TickerTotal + ws.Cells(i, 7).Value
              
                
                ' Print the Ticker Name in the Summary Table
                ws.Range("I" & SummaryRow).Value = TickerName
                ' Print the Ticker Name in the Summary Table
                ws.Range("J" & SummaryRow).Value = PriceChange
                
                
                ' Change Columns colours to GREEN if increase and RED for decrease
                If (PriceChange > 0) Then
                    'Green if there is an increase
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                ElseIf (PriceChange <= 0) Then
                    'Red if there was a decrease
                    ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table
                ws.Range("K" & SummaryRow).Value = (CStr(PercentChange) & "%")
                ' Print the Ticker Name in the Summary Table
                ws.Range("L" & SummaryRow).Value = TickerTotal
                
                ' Add 1 to the summary table row count
                SummaryRow = SummaryRow + 1
                ' Reset PriceChange
                PriceChange = 0
                'next Ticker's Open_Price
                OpenPrice = ws.Cells(i + 1, 3).Value
                End If
     Next i
     Next ws
End Sub


