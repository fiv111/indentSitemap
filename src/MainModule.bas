Attribute VB_Name = "MainModule"
Sub indentURL()
  curtSRow = Selection.Row
  curtLRow = Cells(Rows.Count, Selection.Column).End(xlUp).Row
  'curtSCol = startCell.Column
  'curtLCol = Cells(curtLRow, Columns.Count).End(xlToLeft).Column
  maxCol = 0
  Set target = Selection
  
  ' sort
  target.Sort Key1:=target, Order1:=xlAscending, Header:=xlNo
  
  ' clearContents
  Range(target.Offset(0, 1).Address, Cells(target.Offset(0, 1).Row, target.Offset(0, 1).End(xlToRight).Column)).ClearContents
  
  Call stopCalculate
  
  For i = curtSRow To curtLRow
    Set c = Cells(i, 1)
    
    If c.Value Like "*http*" Then
      u = Split(c.Value, "//")(1)
    Else
      u = c.Value
    End If
    
    urls = Split(u, "/")
    
    For j = 0 To UBound(urls)
      Set cc = Cells(i, j + 2)
      
      If j < UBound(urls) Then
        cc.Value = urls(j) & "/"
      Else
        cc.Value = urls(j)
      End If
      
      Set cc = Nothing
    Next j
    
    
    ' maxCol
    'If maxCol < Cells(i, Columns.Count).End(xlToLeft).Column Then
    '  maxCol = Cells(i, Columns.Count).End(xlToLeft).Column
    'End If
  Next i
  
  Call format_

  Call startCalculate
  Set target = Nothing
End Sub


Private Sub format_()
  For i = Range("B1").CurrentRegion.Rows.Count To 2 Step -1
    For j = Cells(i, Columns.Count).End(xlToLeft).Column To 2 Step -1
      Set c = Cells(i, j)
      'Debug.Print c.Value
        If c.Offset(-1, -1).Value = c.Offset(0, -1).Value And c.Value = c.Offset(-1, 0).Value Then
          c.Value = ""
          c.Interior.ColorIndex = 15
        End If
      
      Set c = Nothing
    Next j
  Next i
  
  Range("B:B").RemoveDuplicates Columns:=1, Header:=xlNo
  Range("B2:B" & Range("B1").CurrentRegion.Rows.Count).Interior.ColorIndex = 15
End Sub


' 自動更新停止
Sub stopCalculate()
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual
End Sub


' 自動更新有効
Sub startCalculate()
  Application.ScreenUpdating = True
  ActiveSheet.EnableCalculation = True
  Application.Calculation = xlCalculationAutomatic
End Sub


Private Sub lining()
    Dim myRng As Range
    Dim c As Range
    Dim Flag As String
    Set myRng = Selection

    Dim i As Long, S As Long, E As Long
    S = Selection(1).Column
    E = Selection(Selection.Count).Column

    For i = Selection(1).Row To Selection(Selection.Count).Row
        Range(Cells(i, S), Cells(i, E)).Select
        Flag = 0
        For Each c In Selection

            If Flag = 0 Or c.Value <> "" Then
                c.Borders(xlEdgeLeft).LineStyle = xlContinuous
            End If

            If c.Value <> "" Or Flag = 1 Then
                c.Borders(xlEdgeTop).LineStyle = xlContinuous
                Flag = 1
            End If
           
        Next c
    Next i
   
    myRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    myRng.Borders(xlEdgeTop).LineStyle = xlContinuous
    myRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    myRng.Borders(xlEdgeRight).LineStyle = xlContinuous
End Sub

