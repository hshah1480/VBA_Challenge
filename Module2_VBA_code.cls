VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Sub stock_analysis()

'set dimension

Dim rowcount As Long
Dim columncount As Long
Dim start As Long

Dim ticker As String

Dim openrate As Double
Dim closerate As Double

Dim yearchange As Double
Dim percentchange As Double
Dim totalvolume As Double

Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatest_totalvolumn As Double

Dim ws As Worksheet


For Each ws In Worksheets

      'set title of the rows
      
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Volume"
      
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      
      'set which row to start
      start = 2
      
      'to find row number of the last row with data
      rowcount = ws.Cells(Rows.Count, "a").End(xlUp).Row
      
      'MsgBox (rowcount)
      
      
      For columncount = 2 To rowcount
            
            ticker = ws.Cells(columncount, 1).Value
            
        If ticker <> ws.Cells(columncount + 1, 1).Value Then
               
               totalvolume = totalvolume + ws.Cells(columncount, 7)
               closerate = ws.Cells(columncount, 6).Value
               yearchange = closerate - openrate
               percentchange = yearchange / openrate
               
               ws.Cells(start, 9).Value = ticker
               ws.Cells(start, 10).Value = yearchange
               ws.Cells(start, 11).Value = percentchange
               ws.Cells(start, 12).Value = totalvolume
               
                  If yearchange > 0 Then
                        ws.Cells(start, 10).Interior.ColorIndex = 4
                    ElseIf yearchange < 0 Then
                        ws.Cells(start, 10).Interior.ColorIndex = 3
                 End If
               
               start = start + 1
               totalvolume = 0
               
            Else
                     totalvolume = totalvolume + ws.Cells(columncount, 7).Value
                 If ticker <> ws.Cells(columncount - 1, 1).Value Then
                    openrate = ws.Cells(start, 3).Value
                 End If
              
            
        End If
            
            
       Next columncount

      'take max and min
      
      ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
      ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
      ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))
      
      increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
      decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
      volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)

     ws.Range("P2") = ws.Cells(increase_number + 1, 9)
     ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
     ws.Range("P4") = ws.Cells(volume_number + 1, 9)







Next ws

End Sub

