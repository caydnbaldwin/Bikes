Option Explicit

Sub Bikes()

'label variables
Dim NOD As String 'name of date
Dim L As String 'location
Dim Q As Integer 'quantity
Dim Rev As Currency 'revenue
Dim Rat As Integer 'rating
Dim inputRow As Integer 'travel to input column
Dim outputRow As Integer 'travel to output column
Dim SS As String 'summary sentence
Dim OTQ As Range, OAR As Range, OTR As Range
Dim PTQ As Range, PAR As Range, PTR As Range
Dim STQ As Range, SAR As Range, STR As Range
Dim GTTQ As Range, GTAR As Range, GTTR As Range
Dim ONOR As Integer 'count orem ratings
Dim PNOR As Integer 'count provo ratings
Dim SNOR As Integer 'count springville ratings
Dim GTNOR As Integer 'count grand total ratings

'locate starting point
Sheets("Bikes").Activate
Range("B14").Activate
outputRow = 13
ONOR = 0
PNOR = 0
SNOR = 0
GTNOR = 0

'clear contents
Range("E15:G18").ClearContents
Range("I13:I10000").ClearContents

'begin loop
Do Until IsEmpty(ActiveCell)
  DoEvents
  
  'input variables
  NOD = ActiveCell.Value
  L = ActiveCell.Offset(1, 0).Value
  Q = ActiveCell.Offset(2, 0).Value
  Rev = ActiveCell.Offset(3, 0).Value
  Rat = ActiveCell.Offset(4, 0).Value
  
    'correct input
    L = Left(LCase(L), 3)
    If L = "ore" Then
        L = "Orem"
      ElseIf L = "pro" Then
        L = "Provo"
      ElseIf L = "spr" Then
        L = "Springville"
    End If
        
  'output SS
  SS = NOD & " sold " & Q & " bikes at the " & L & " office for a total of $" & Rev & "."
  Cells(outputRow, "I").Value = SS
  
  'summary statistics table
  'locate
  Set OTQ = Range("E15")
  Set OAR = Range("F15")
  Set OTR = Range("G15")
  Set PTQ = Range("E16")
  Set PAR = Range("F16")
  Set PTR = Range("G16")
  Set STQ = Range("E17")
  Set SAR = Range("F17")
  Set STR = Range("G17")
  
  'format
  OAR.NumberFormat = "0.00"
  PAR.NumberFormat = "0.00"
  SAR.NumberFormat = "0.00"

  'output by location
  If L = "Orem" Then
      OTQ = OTQ + Q
      OAR = OAR + Rat
      OTR = OTR + Rev
      ONOR = ONOR + 1 'for average calculation
    ElseIf L = "Provo" Then
      PTQ = PTQ + Q
      PAR = PAR + Rat
      PTR = PTR + Rev
      PNOR = PNOR + 1 'for average calculation
    ElseIf L = "Springville" Then
      STQ = STQ + Q
      SAR = SAR + Rat
      STR = STR + Rev
      SNOR = SNOR + 1 'for average calculation
  End If

  'travel to next input
  ActiveCell.End(xlDown).End(xlDown).Activate
  
  'increment
  outputRow = outputRow + 1
  GTNOR = GTNOR + 1
  
Loop

'output average ratings
OAR = OAR / ONOR
PAR = PAR / PNOR
SAR = SAR / SNOR

'grand totals
  'locate
  Set GTTQ = Range("E18")
  Set GTAR = Range("F18")
  Set GTTR = Range("G18")
  GTNOR = Range("F15:F17").Count
  
  'format
  GTAR.NumberFormat = "0.00"
  
  'output by variable
  GTTQ = OTQ + PTQ + STQ
  GTAR = (OAR + PAR + SAR) / GTNOR
  GTTR = OTR + PTR + STR

'back to top
Range("A1").Activate
Beep

End Sub
