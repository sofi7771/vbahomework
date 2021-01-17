Attribute VB_Name = "Module1"




'question 1 Create a script that will loop through all the stocks for one year and output the following information.
'The ticker symbol.

Sub ticker()

Dim tickername As String
'set a row counter
Dim tickervolume As Double
         tickervolume = 0
         Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        Dim open_price As Double
               'initial open price
               open_price = Cells(2, 3).Value
        
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

 For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
             tickername = Cells(i, 1).Value
                Range("I" & summary_ticker_row).Value = tickername
                Cells(1, 9).Value = "Tickersymbol"
              
 
'question 2 Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

             
               
        
                Dim close_price As Double
                Dim yearly_change As Double
                
               
               
               
               close_price = Cells(i, 6).Value
               
                yearly_change = (close_price - open_price)
              
              Range("J" & summary_ticker_row).Value = yearly_change
               Cells(1, 10).Value = "Yearly Change"
              
              
              
 'question 3 The percent change from opening price at the beginning of a given year to the closing price at the end of that year.


                Dim percent_change As Double

              'since 0/0 gives mathematical error, that has to be avoided:
              
               If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If


              Cells(1, 11).Value = "Percent Change"
            Range("K" & summary_ticker_row).Value = percent_change
            Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   

             
            

'question 4 The total stock volume of the stock.

 

      tickervolume = tickervolume + Cells(i, 7).Value
      Cells(1, 12).Value = "Total Stock Volume"
      Range("L" & summary_ticker_row).Value = tickervolume
      tickervolume = 0
      
              summary_ticker_row = summary_ticker_row + 1

      open_price = Cells(i + 1, 3).Value
Else
              
               'Add the volume of trade
              tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next



'question 5 You should also have conditional formatting that will highlight positive change in green and negative change in red.


     
lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    


 For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i



'bonus question


        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"



 For i = 2 To lastrow_summary_table
            'Find the maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub










       

           

