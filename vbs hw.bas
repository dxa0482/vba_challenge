Attribute VB_Name = "Module1"
Sub stonks()

      ' Set an initial variable for holding the brand name
      Dim Stock As String
    
      ' Set an initial variable for holding the total Volume per stock
      Dim stockVolume As Double
      stockVolume = 0
      
      
      ' Keep track of stock open and close
      Dim stockOpen As Double
      Dim StockClose As Double
      Dim yearlyChange As Double
      Dim percentageChange As Double
      
      
      stockOpen = Cells(2, 3).Value
      
    
      ' Keep track of the location for each credit card brand in the summary table
      Dim SummaryTable_Row As Integer
      summaryTableRow = 2
      
      
      Dim lastRow As Long
      'Find the last non-blank cell in column A(1)
      lastRow = Cells(Rows.Count, 1).End(xlUp).Row
     
    
      ' Loop through all stock prices
      For i = 2 To lastRow
    
            ' Check if we are still within the same stock, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              ' Set the Stock name
              Stock = Cells(i, 1).Value
        
              ' Add to the Total Volume
              stockVolume = stockVolume + Cells(i, 7).Value
              
              ' Get the close
              StockClose = Cells(i, 6).Value
              yearlyChange = StockClose - stockOpen
              
             
              ' percent change
                If stockOpen > 0 Then
                    percentChange = yearlyChange / stockOpen
                Else
                    percentChange = -1
                End If
        
              ' Print the Credit Card Brand in the Summary Table
              Cells(summaryTableRow, 9).Value = Stock
              
        
              ' Print the Total Volume to the Summary Table
              Cells(summaryTableRow, 12).Value = stockVolume
              Cells(summaryTableRow, 10).Value = yearlyChange
              Cells(summaryTableRow, 11).Value = percentageChange
              
              
              ' Color Code
              If yearlyChange < 0 Then
               Cells(summaryTableRow, 10).Interior.ColorIndex = 3
               
              Else
                Cells(summaryTableRow, 10).Interior.ColorIndex = 4
               
              End If
              
        
               ' RESETS
    
            
               ' Add one to the summary table row
                  summaryTableRow = summaryTableRow + 1
                  
               ' Reset the Brand Total
                  stockVolume = 0
                  
                  percentage = 0
                  
               ' Reset stock open
               stockOpen = Cells(i + 1, 3).Value
            
                ' If the cell immediately following a row is the same stock...
                Else
            
                  ' Add to the Total Volume
                  stockVolume = stockVolume + Cells(i, 7).Value
        
            End If

  Next i


End Sub

