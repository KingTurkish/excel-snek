Attribute VB_Name = "Main"
Public rinc As Integer, cinc As Integer
Dim r() As Integer, c() As Integer, t As Double
Public game As Boolean
Dim apples As Integer, gameOver As Boolean

Dim boardColor As Long, snakeBodyColor As Long, snakeheadColor As Long, appleColor As Long
Sub SetColor()
    boardColor = RGB(255, 246, 211)
    
    snakeBodyColor = RGB(249, 168, 117)
    snakeheadColor = RGB(235, 107, 111)
    
    appleColor = RGB(124, 63, 88)
End Sub
Sub SetScore()
    Range("B27").Value = "Score: " & apples
End Sub
Sub StartGame()
    game = True
    gameOver = False
    apples = 0
    SetScore
    SetColor
    CreateBoard
    ChangeButtonText
    CreateSnake
    SpawnApple
    ShowSnake
    bindKeys
    Update

End Sub
Sub ResetGame()
    apples = 0
    SetScore
    SetColor
    CreateBoard
    CreateSnake
    ShowSnake
    bindKeys
    game = False
End Sub
Sub TogglePause()
    If gameOver <> False Then Exit Sub
    game = Not game
    Update
    ChangeButtonText
End Sub
Sub ChangeButtonText()
    If game = True Then
        ActiveSheet.Shapes("Toggle").TextFrame.Characters.Text = "Pause"
    Else
        ActiveSheet.Shapes("Toggle").TextFrame.Characters.Text = "Resume"
    End If
End Sub
Sub CreateBoard()
    'Board Setup
    Range("B2:Z26").Clear
    Range("B2:Z26").Interior.color = boardColor
    Range("B2:Z26").Borders(xlEdgeBottom).color = appleColor
    Range("B2:Z26").Borders(xlEdgeLeft).color = appleColor
    Range("B2:Z26").Borders(xlEdgeTop).color = appleColor
    Range("B2:Z26").Borders(xlEdgeRight).color = appleColor
    
End Sub

Sub CreateSnake()
    ReDim r(2)
    ReDim c(2)
    r(0) = 20: r(1) = 21: r(2) = 22
    c(0) = 14: c(1) = 14: c(2) = 14
    rinc = -1: cinc = 0
End Sub
Sub ShowSnake()
    For i = UBound(r) To 1 Step -1
        Cells(r(i), c(i)).Interior.color = snakeBodyColor
    Next i
    Cells(r(0), c(0)).Interior.color = snakeheadColor
End Sub

Sub MoveSnake()
    If rinc <> 0 Or cinc <> 0 Then
        SetScore
        'get last cell to remove from snake
        tail = UBound(r)
        Cells(r(tail), c(tail)).Interior.color = boardColor
        
        'move snake body
        For i = tail To 1 Step -1
            r(i) = r(i - 1)
            c(i) = c(i - 1)
        Next i
        'move snake head
        r(0) = r(0) + rinc
        c(0) = c(0) + cinc
        'check snake movement
        If Cells(r(0), c(0)).Interior.color = appleColor Then
            'pick apple
            apples = apples + 1
            ReDim Preserve r(UBound(r) + 1)
            ReDim Preserve c(UBound(c) + 1)
            r(UBound(r)) = r(UBound(r) - 1)
            c(UBound(c)) = c(UBound(c) - 1)
            SpawnApple
        ElseIf Cells(r(0), c(0)).Interior.color <> boardColor Then
            game = False
            gameOver = True
            MsgBox "Game Over"
            Exit Sub
        End If
                

        ShowSnake
    End If
End Sub

Sub SpawnApple()
    Randomize
    arow = Int(Rnd * 23) + 2
    acol = Int(Rnd * 23) + 2
    If Cells(arow, acol).Interior.color = snakeBodyColor Then
     SpawnApple
     Exit Sub
    End If
    Cells(arow, acol).Interior.color = appleColor
End Sub

Sub Update()
    If game <> True Then Exit Sub
    MoveSnake
    'StartTimer
    Application.OnTime DateAdd("s", 1, Now), "Update"
End Sub

