Attribute VB_Name = "Module1"
Sub stock_analysis()

' This macro analyzes stock data and creates an annual summary on a new sheet.

  ' Set an initial variable for holding the ticker symbol of each stock
  Dim Ticker As String

  ' Set an initial variable for holding the total volume traded by ticker symbol
  Dim Volume_Total As Double
  Volume_Total = 0
  
  ' Set initial variable for holding the opening price by ticker symbol, and set first open
  Dim First_Open As Double
  
  ' Set an initial variable for holding the closing price by ticker symbol
  Dim Last_Close As Double
  
  'Set an initial variable for holding the yearly change by ticker symbol
  Dim Yearly_Change As Double
  
  'Set an initial variable for holding the yearly percent change by ticker symbol
  Dim Yearly_Percent_Change As Double
  
  'Set an initial variable for holding the greatest % increase
  Dim Greatest_Percent_Increase As Double
  Greatest_Percent_Increase = 0
  
  'Set an initial variable for holding the greatest % decrease
  Dim Greatest_Percent_Decrease As Double
  Greatest_Percent_Decrease = 0
  
  'Set an initial variable for holding the greatest total volume
  Dim Greatest_Volume_Total As Double
  Greatest_Volume_Total = 0
  
  ' Insert new sheet for summary table to the front of the workbook
  ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1), Type:=xlWorksheet).Name = "SummaryTable"
  
  
  ' Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 8
  
  ' Detemine the Last Worksheet
  Dim Last_WS As Integer
  LastWS = Application.Sheets.Count
  
 
  For w = 2 To LastWS
  
    ' Creates column headers for summary table
    Dim headers() As Variant
    headers() = Array("Ticker", "Annual Change ($)", "Annual Change (%)", "Total Volume")
    
    With ActiveWorkbook.Sheets(1)
    For c = LBound(headers()) To UBound(headers())
      .Cells(7, 2 + (w - 2) * 5 + c).Value = headers(c)
      Next c
    End With
    
    ' Creates row and column headers for challenge summary table
    Dim Summary_Table_Name As String
    Summary_Table_Name = ActiveWorkbook.Sheets(w).Name + " - Summary"
    
    ActiveWorkbook.Sheets(1).Cells(1, 2 + (w - 2) * 5) = Summary_Table_Name
    ActiveWorkbook.Sheets(1).Cells(2, 3 + (w - 2) * 5) = "Ticker"
    ActiveWorkbook.Sheets(1).Cells(2, 4 + (w - 2) * 5) = "Value"
    ActiveWorkbook.Sheets(1).Cells(3, 2 + (w - 2) * 5) = "Greatest % Increase"
    ActiveWorkbook.Sheets(1).Cells(4, 2 + (w - 2) * 5) = "Greatest % Decrease"
    ActiveWorkbook.Sheets(1).Cells(5, 2 + (w - 2) * 5) = "Greatest Total Volume"
        
     ' Determine the Last Row in each worksheet
     LastRowData = ActiveWorkbook.Sheets(w).Cells(Rows.Count, 1).End(xlUp).Row
     
     First_Open = ActiveWorkbook.Sheets(w).Cells(2, 3).Value
      
        ' Loop through all stock data
        For i = 2 To LastRowData
      
          ' Check if we are still within the same ticker symbol, if it is not...
          If ActiveWorkbook.Sheets(w).Cells(i + 1, 1).Value <> ActiveWorkbook.Sheets(w).Cells(i, 1).Value Then
            
            ' Set the ticker symbol
            Ticker = ActiveWorkbook.Sheets(w).Cells(i, 1).Value
            
            ' Set yearly change
            Last_Close = ActiveWorkbook.Sheets(w).Cells(i, 6).Value
            Yearly_Change = Last_Close - First_Open
            
            ' Set yearly percent change
            If First_Open = 0 Then
            Yearly_Percent_Change = 0
            Else
            Yearly_Percent_Change = (Last_Close - First_Open) / First_Open
            End If
            
            ' Set yearly percent change to greatest increase if bigger than prior, or to greatest decrease if bigger than prior
            If Yearly_Percent_Change >= Greatest_Percent_Increase Then
              Greatest_Percent_Increase = Yearly_Percent_Change
            ElseIf Yearly_Percent_Change <= Greatest_Percent_Decrease Then
              Greatest_Percent_Decrease = Yearly_Percent_Change
              End If
              
            ' Drops ticker and volume into summary tab if % increase or decrease is largest
            If Greatest_Percent_Increase = Yearly_Percent_Change Then
              ActiveWorkbook.Sheets(1).Cells(3, 3 + (w - 2) * 5) = Ticker
              End If
              
            If Greatest_Percent_Increase = Yearly_Percent_Change Then
              ActiveWorkbook.Sheets(1).Cells(3, 4 + (w - 2) * 5) = Format(Greatest_Percent_Increase, "0%")
              End If
              
            If Greatest_Percent_Decrease = Yearly_Percent_Change Then
              ActiveWorkbook.Sheets(1).Cells(4, 3 + (w - 2) * 5) = Ticker
              End If
            
            If Greatest_Percent_Decrease = Yearly_Percent_Change Then
              ActiveWorkbook.Sheets(1).Cells(4, 4 + (w - 2) * 5) = Format(Greatest_Percent_Decrease, "0%")
              End If
              
            ' Add to the volume total
            Volume_Total = Volume_Total + ActiveWorkbook.Sheets(w).Cells(i, 7).Value
            
            ' Set total volume to greatest if bigger than prior
            If Volume_Total >= Greatest_Volume_Total Then
              Greatest_Volume_Total = Volume_Total
            End If
      
            ' Drops ticker and volume into summary tab if % volume total largest
            If Greatest_Volume_Total = Volume_Total Then
              ActiveWorkbook.Sheets(1).Cells(5, 3 + (w - 2) * 5) = Ticker
              End If
              
            If Greatest_Volume_Total = Volume_Total Then
              ActiveWorkbook.Sheets(1).Cells(5, 4 + (w - 2) * 5) = Format(Volume_Total, "#,###")
              End If
            
            ' Print the ticker symbol in the Summary Table
            ActiveWorkbook.Sheets(1).Cells(Summary_Table_Row, 2 + (w - 2) * 5).Value = Ticker
            
            ' Print the yearly change in the Summary Table
            ActiveWorkbook.Sheets(1).Cells(Summary_Table_Row, 3 + (w - 2) * 5).Value = Yearly_Change
            
            ' Print the yearly percent change in the Summary Table
            ActiveWorkbook.Sheets(1).Cells(Summary_Table_Row, 4 + (w - 2) * 5).Value = Format(Yearly_Percent_Change, "0%")
      
            ' Print the total volume to the Summary Table
            ActiveWorkbook.Sheets(1).Cells(Summary_Table_Row, 5 + (w - 2) * 5).Value = Format(Volume_Total, "#,###")
      
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            First_Open = ActiveWorkbook.Sheets(w).Cells(i + 1, 3).Value
            
            ' Reset the Brand Total
            Volume_Total = 0
      
          ' If the cell immediately following a row is the same brand...
          Else
      
            ' Add to the Brand Total
            Volume_Total = Volume_Total + ActiveWorkbook.Sheets(w).Cells(i, 7).Value
      
          End If
      
        Next i
        
        Summary_Table_Row = 8
        Greatest_Percent_Increase = 0
        Greatest_Percent_Decrease = 0
        Greatest_Volume_Total = 0
        
        'Applies conditional formatting to annual change column of summary table
  
        Dim ColorCodeRange As Range
        Set ColorCodeRange = ActiveWorkbook.Sheets(1).Range(Cells(8, 3 + (w - 2) * 5), Cells(8, 3 + (w - 2) * 5).End(xlDown))
        
        With ColorCodeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=" & 0)
          .Interior.Color = rgbLimeGreen
        End With
          
        With ColorCodeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & 0)
          .Interior.Color = rgbRed
        End With
        
    Next w
      
  
  'Auto-fits column sizes for summary table sheet
  ActiveWorkbook.Sheets(1).Cells.EntireColumn.AutoFit
  
  
End Sub
