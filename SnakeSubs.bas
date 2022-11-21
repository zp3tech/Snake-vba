Option Explicit
Public Head As Range
Public SnakeLength As Integer
Public Direction As Integer  '1=UP 2=RIGHT 3=DOWN 4=LEFT

Sub ReduceBoardValues()
  Dim cv as Range
  For Each cv In Range("B2:Z26")
    If cv.Value <> 0 And cv.Value <> "O" Then
      cv.Value = cv.Value - 1
      If cv.Value = 0 Then
        cv.ClearContents
        With cv.Interior
          .Pattern = xlNone
          .TintAndShade = 0
          .PatternTintAndShade = 0
        End With
        cv.Font.ThemeColor = xlThemeColorAccent6  'reset font color
      End If
    End If
  Next cv
End Sub

Sub BoardSetup()
  'redraw borders in black
  Range("A1:AA27").Select
  With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With
  Range("B2:Z26").Select
  Selection.ClearContents
  Selection.Font.ThemeColor = xlThemeColorAccent6
  With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With

  Range("I31").Select
End Sub

Sub StartGame()
  Call BoardSetup
  'sets initial game conditions
  SnakeLength = 4
  Set Head = Range("B9")
  Head.Interior.Color = 5287936  'light green
  Range("B30").Value = 0   'start scoreboard at zero
  Direction = 2   'start snake moving to the right

  Call CreateFood
  Call ContinuousMove

  MsgBox ("Game over!" & vbCrLf & "Score: " & Range("B30").Value)
End Sub

Sub ContinuousMove()
  'UP
  If Direction = 1 Then
    'check if head is in wall or body of snake
    Do While Head.Value = ""
      Call TimeToAct
      If Direction <> 1 Then
        Call ContinuousMove
        Exit Do
      End If
      Head.Value = SnakeLength
      Call ReduceBoardValues
      Set Head = Head.Offset(-1, 0)
      Head.Interior.Color = 5287936
      Call EatFood
    Loop
  End If 

  'RIGHT
  If Direction = 2 Then
    'check if head is in wall or body of snake
    Do While Head.Value = ""
      Call TimeToAct
      If Direction <> 2 Then
        Call ContinuousMove
        Exit Do
      End If
      Head.Value = SnakeLength
      Call ReduceBoardValues
      Set Head = Head.Offset(0, 1)
      Head.Interior.Color = 5287936
      Call EatFood
    Loop
  End If 

  'DOWN
  If Direction = 3 Then
    'check if head is in wall or body of snake
    Do While Head.Value = ""
      Call TimeToAct
      If Direction <> 3 Then
        Call ContinuousMove
        Exit Do
      End If
      Head.Value = SnakeLength
      Call ReduceBoardValues
      Set Head = Head.Offset(1, 0)
      Head.Interior.Color = 5287936
      Call EatFood
    Loop
  End If 

  'LEFT
  If Direction = 4 Then
    'check if head is in wall or body of snake
    Do While Head.Value = ""
      Call TimeToAct
      If Direction <> 4 Then
        Call ContinuousMove
        Exit Do
      End If
      Head.Value = SnakeLength
      Call ReduceBoardValues
      Set Head = Head.Offset(0, -1)
      Head.Interior.Color = 5287936
      Call EatFood
    Loop
  End If 
End Sub

Sub EatFood()
  If Head.Value = "O" Then 
    SnakeLength = SnakeLength + 2
    Head.Value = ""  'removes the "O" from the board
    Head.Interior.ThemeColor = xlThemeColorAccent6
    Head.Font.ThemeColor = xlThemeColorAccent6
    Range("B30").Value = Range("B30").Value + 1  'increment scoreboard
    Call CreateFood
  End if
End Sub

Sub CreateFood()
  Dim RowRandom As Integer
  Dim ColRandom As Integer
  RowRandom = Int((26-2+1) * Rnd + 2)
  ColRandom = Int((26-2+1) * Rnd + 2)
  If ThisWorkbook.Worksheets(1).Cells(RowRandom, ColRandom).Value <> "" Then
    Call CreateFood
    Exit Sub
  End If
  ThisWorkbook.Worksheets(1).Cells(RowRandom, ColRandom).Value = "O"   'capital letter "oh"
End Sub

Sub TimeToAct()
  Dim NowTick As Long
  Dim EndTick As Long
  NowTick = 0
  EndTick = 500
  Do Until NowTick > EndTick
    DoEvents
    If ActiveCell.Value = "UP" And Direction <> 3 Then
      Direction = 1 
      Range("I31").Select
      Exit Sub
    End If
    If ActiveCell.Value = "RIGHT" And Direction <> 4 Then
      Direction = 2 
      Range("I31").Select
      Exit Sub
    End If
    If ActiveCell.Value = "DOWN" And Direction <> 1 Then
      Direction = 3 
      Range("I31").Select
      Exit Sub
    End If
    If ActiveCell.Value = "LEFT" And Direction <> 2 Then
      Direction = 4 
      Range("I31").Select
      Exit Sub
    End If
    NowTick = NowTick + 1
  Loop
  Range("I31").Select
End Sub