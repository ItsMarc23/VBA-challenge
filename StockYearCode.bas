Attribute VB_Name = "Module1"
Sub abcTest()
        Dim ws As Worksheet
        Dim Stock_Name As String
        Dim Stock_Total As Double
        Dim Stock_Year As Double
        Dim Stock_Percent As Double
        Dim Summary_Table_Row As Integer
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim Year_Change As Double
        Dim Percent_Change As String
        Dim Next_Ticker As Long
        Dim Max_Value As Double
        Dim Low_Value As Double
        Dim Counter As Integer
        Dim Max_Stock As LongLong
        Dim St1 As String
        Dim St2 As String
        Dim St3 As String
        Dim Ticker_Open As LongLong
        
    For Each ws In Worksheets
       Summary_Table_Row = 2
        Next_Ticker = 2
        Counter = 2
        Ticker_Open = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Range("H1").EntireColumn.Insert
    
        ws.Cells(1, 8).Value = "Ticker"
        
        ws.Range("I1").EntireColumn.Insert
        
        ws.Cells(1, 9).Value = "YearlyChange"
            
        ws.Range("J1").EntireColumn.Insert
        
        ws.Cells(1, 10).Value = "Percent Change"
        
        ws.Range("K1").EntireColumn.Insert
     
        ws.Cells(1, 11).Value = "Total Stock Volume"
        
        ws.Range("O1").EntireColumn.Insert
        
        ws.Cells(1, 15).Value = "Ticker"
        
        ws.Range("P1").EntireColumn.Insert
        
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 14).Value = "Greates % Increase"
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
           
        For i = 2 To LastRow
        
        
        
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                OpenPrice = ws.Cells(Ticker_Open, 3).Value
                
               Stock_Name = ws.Cells(i, 1).Value
                
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
        
                ws.Range("H" & Summary_Table_Row).Value = Stock_Name

                ws.Range("K" & Summary_Table_Row).Value = Stock_Total
            
                Stock_Total = 0
                        
               ClosePrice = ws.Cells(i, 6).Value
               
               
               Year_Change = (ClosePrice - OpenPrice)
               
               ws.Range("I" & Next_Ticker).Value = Year_Change
               
               If ws.Range("I" & Next_Ticker) > 0 Then
               ws.Range("I" & Next_Ticker).Interior.ColorIndex = 4

            Else
               ws.Range("I" & Next_Ticker).Interior.ColorIndex = 3

            End If
            
              Percent_Change = FormatPercent((ClosePrice - OpenPrice) / OpenPrice)
               
               ws.Range("J" & Next_Ticker).Value = Percent_Change
            
                Next_Ticker = Next_Ticker + 1
                
                 Summary_Table_Row = Summary_Table_Row + 1
                 
                Ticker_Open = i + 1
                
         Else
           
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
             
                 
            End If
        
     
               
            Next i
             For x = 2 To LastRow
                   If ws.Cells(x, 10).Value > Max_Value Then
                   Max_Value = ws.Cells(x, 10).Value
                   St1 = ws.Cells(x, 8).Value
                   ws.Range("O" & Counter).Value = St1
                   
                   End If
                    
                 Next x
                
                ws.Range("P" & Counter).Value = FormatPercent(Max_Value)
                
                 For Z = 2 To LastRow
                   If ws.Range("J" & Z).Value < Low_Value Then
                   Low_Value = ws.Range("J" & Z).Value
                    St2 = ws.Cells(Z, 8).Value
                   ws.Range("O" & Counter + 1).Value = St2
                   
                   End If
                    
                 Next Z
                    ws.Range("P" & Counter + 1).Value = FormatPercent(Low_Value)
                     
                     
                   For y = 2 To LastRow
                  If ws.Cells(y, 11).Value > Max_Stock Then
                   Max_Stock = ws.Cells(y, 11).Value
                     St3 = ws.Cells(y, 8).Value
                   ws.Range("O" & Counter + 2).Value = St3
                   
                   End If
                    ws.Range("P" & Counter + 2).Value = Max_Stock
                
                 Next y

    Next ws
    
            
End Sub

