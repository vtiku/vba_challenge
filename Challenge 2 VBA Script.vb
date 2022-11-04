
Sub Ticker_Summary()

 ' Define worksheet
 Dim Current As Worksheet
 For Each Current In Worksheets
  ' Set an initial variable for holding the Ticker Symbol
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the total volume per Ticker Symbol
  Dim Total_Volume As LongLong
  Total_Volume = 0

  ' Define Opening Price
  Dim Open_Price As Double
  Open_Price = 0

' Define Closing Price
  Dim Close_Price As Double
  Close_Price = 0
  
  'Define Outputs / Formulas
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  
  'set up last row definition
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'To fix overflow error (from google)
On Error Resume Next


'label headers
   Cells(1, 9).Value = "Ticker"
   Cells(1, 10).Value = "Yearly Change"
   Cells(1, 11).Value = "Percent Change"
   Cells(1, 12).Value = "Total Stock Volume"

 sheetName = InputBox("What data sheet would you like to run the calculations on?")
Sheets(sheetName).Activate
      
  ' Loop through all Dates
  For i = 2 To lastrow
   
   
    
    ' Check if we are still within the same Ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Symbol
      Ticker_Name = Cells(i, 1).Value
     
      '(Yearly/% Change inputs
          Close_Price = Cells(i, 6).Value
      
      'Yearly/% Change formulas
      Yearly_Change = Close_Price - Open_Price
      Percent_Change = (Close_Price - Open_Price) / Close_Price
            
      
                
      ' Add to the Total
      Total_Volume = Total_Volume + Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

    'Print Yearly Change to the Summary Table
    Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      'Print Percent Change to the Summary Table
    Range("K" & Summary_Table_Row).Value = Percent_Change
           
      ' Print the Brand Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Total_Volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else


      ' Add to the  Total volume
      Total_Volume = Total_Volume + Cells(i, 7).Value
    'outside conditional for open price
    
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
             '(Yearly/% Change inputs
                Open_Price = Cells(i, 3).Value
        
    
      
            End If
            
             
        
        
        
    
    End If

  Next i
        For j = 2 To lastrow
        
        If Cells(j, 11) > 0 Then
            
            Cells(j, 11).Interior.Color = vbGreen
            
        Else
        
            Cells(j, 11).Interior.Color = vbRed
            
        End If
        
    Next j
                
    'format columns
    Columns("K").NumberFormat = "0.00%"
   
     
 Next

 
End Sub