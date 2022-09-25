Attribute VB_Name = "Module1"
Sub Stocks()


Dim sheets As Worksheet

    For Each sheets In ThisWorkbook.Worksheets
    
        sheets.Activate
         
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
 

        
        Dim strt      As Double
        Dim ychg      As Double
        Dim pchg      As Double
        Dim rvolume   As Double
        Dim volume    As Double
        Dim crvolume  As Double
        Dim rgpchg    As Double
        Dim rlpchg    As Double
        Dim nname     As Double
        Dim j         As Double
        Dim n         As Double
        Dim Num_Row   As Double
        Dim Nnum_Row  As Double
        
        strt = 0
        ychg = 0
        pchg = 0
        volume = 0
        rvolume = 0
        crvolume = 0
        rgpchg = 0
        rlpchg = 0
        nname = 0
        Num_Row = 0
        Nnum_Row = 0
        
                

        Num_Row = Cells(Rows.Count, 1).End(xlUp).Row
        Nnum_Row = Num_Row
        
        
        j = 2
        
          
        For i = 2 To Num_Row
            If strt = 0 Then
               strt = Cells(i, 3).Value
            End If
            
             
            
            If Cells(i, "A").Value = Cells(i + 1, "A").Value Then
                volume = volume + Cells(i, 7).Value
            
            ElseIf Cells(i, "A").Value <> Cells(i + 1, "A").Value Then
                
                ychg = (Cells(i, 6) - strt)
                pchg = (ychg / strt)
                dpchg = (pchg * 2)
                
                
                volume = volume + Cells(i, 7).Value
                Cells(j, 9).Value = Cells(i, "A").Value
                Cells(j, 10).Value = ychg
                Cells(j, 11).Value = FormatPercent(pchg, 2)
                Cells(j, 12).Value = volume

                
                
                If Cells(j, 10).Value > 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                ElseIf Cells(j, 10).Value < 0 Then
                    Cells(j, 10).Interior.ColorIndex = 3
                End If
                
                strt = 0
                volume = 0
                
                j = j + 1
                                               
            End If
            
             
         
         Next i
          
        sheet_name = ActiveSheet.Name
        
        Worksheets(sheet_name).Columns("J").AutoFit
        Worksheets(sheet_name).Columns("K").AutoFit
        Worksheets(sheet_name).Columns("L").AutoFit
        Worksheets(sheet_name).Columns("O").AutoFit
        
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
                      
        rgpchg = Application.WorksheetFunction.Max(Range("K2:K" & Num_Row))
        rlpchg = Application.WorksheetFunction.Min(Range("K2:K" & Num_Row))
        rvolume = Application.WorksheetFunction.Max(Range("L2:L" & Num_Row))
        
        
                       
        Cells(2, 17).Value = rgpchg
         
        Cells(3, 17).Value = rlpchg
               
        Cells(4, 17).Value = rvolume
        
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        
           
        n = 2
        
        For n = 2 To Nnum_Row
        
        ' If crvolume = 0 Then
        '    crvolume = Cells(4, "Q").Value
        'End If
          
         If Cells(n, "K").Value = Cells(2, 17).Value Then
             Cells(2, 16).Value = Cells(n, "I").Value
         End If
             
         If Cells(n, "K").Value = Cells(3, 17).Value Then
              Cells(3, 16).Value = Cells(n, "I").Value
         End If
         
         If Cells(n, "L").Value = Cells(4, 17) Then
              Cells(4, 16).Value = Cells(n, "I").Value
         End If
                   
                   
          Next n
               
          
          Range("Q4").NumberFormat = "0.00E+00"
          
          Worksheets(sheet_name).Columns("O").AutoFit
                   
          Worksheets(sheet_name).Columns("Q").AutoFit
          
      Next sheets
                
             
       

        

End Sub






