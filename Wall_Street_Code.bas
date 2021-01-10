Attribute VB_Name = "Module14"
Sub AlphabeticalTesting()

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"
Dim Column As Integer
Column = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim Yearly_Change As Double
Yearly_Change = 0

Dim Total_Percent As Double
Total_Percent = 0
Dim Total_Stock As Double
Total_Stock = 0

Dim Close_Price As Double

Dim Open_Price As Double




    For i = 2 To LastRow
    
    
        
                    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'ElseIf Open_Price = 0 Then
            'Total_Percent = 0
            
            ticker = Cells(i, 1).Value
        
            Range("I" & Column).Value = ticker
            
            
            Total_Stock = Total_Stock + Cells(i, 7).Value
        
            Range("L" & Column).Value = Total_Stock
                   
            
            Close_Price = Cells(i, 6).Value
            
            Yearly_Change = Close_Price - Open_Price
        
            Range("J" & Column).Value = Yearly_Change
            
            
            
            
            
            If Open_Price = 0 Then
                Cells(i, 11).Value = Null
            Else
                Cells(i, 11).Value = (Close_Price - Open_Price) / Open_Price
                
                Total_Percent = ((Close_Price - Open_Price) / Open_Price)
            
            Range("K" & Column).Value = Total_Percent
            
            End If
            
           
            
        
                   
            Column = Column + 1
            
            Total_Stock = 0
            
            Open_Price = 0
            
            Close_Price = 0
              
            
        
        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            Open_Price = Cells(i, 3).Value
            
            
        
            
            Yearly_Change = Yearly_Change + Close_Price - Open_Price
            
            
            Total_Stock = Total_Stock + Cells(i, 7).Value
            
        
           
            
        Else
        
            
            
        
            Total_Stock = Total_Stock + Cells(i, 7).Value
        
           
        End If
           
           
    
        
        If Range("J" & Column).Value >= 0 Then
            Range("J" & Column).Interior.ColorIndex = 4
        
        Else
        
        Range("J" & Column).Interior.ColorIndex = 3
        
        End If
        
            
            
            
        
            
        
        Next i
        
             
End Sub



