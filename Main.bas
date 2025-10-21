Attribute VB_Name = "Main"
'======================
' Excel Snek Game - v1.0
'
' Original concept and code by Haronkar Singh (Excel-Snake)
' https://github.com/haronkar/excel-snake
'
' Adapted and improved by Tristan Caldwell (Excel-Snek)
' https://github.com/KingTurkish/Excel-Snek
'
' Lots of help from ChatGPT
' https://chatgpt.com/
'
'======================

Option Explicit

'======================
' GLOBAL VARIABLES
'======================
Public rinc As Integer, cinc As Integer
Dim r() As Integer, c() As Integer
Public game As Boolean
Public running As Boolean, paused As Boolean
Dim apples As Long, extraCoins As Long, applesEaten As Long
Public level As Long, applesForNextLevel As Integer
Public moveInterval As Double, baseSpeed As Double
Public rng As Range

' Colors
Dim boardColor As Long, snakeBodyColor As Long, snakeheadColor As Long
Dim appleColor As Long, wallColor As Long

' Walls
Dim wallRows() As Long, wallCols() As Long
Dim wallCount As Long
Public wallPassActive As Boolean

' Traps (Level 18+)
Dim trapRows() As Long, trapCols() As Long, trapCount As Long
Dim trapFlashing As Boolean
Dim trapFlashEnd As Double
Dim trapFlashRow As Long
Dim trapFlashCol As Long

' Power-ups
Dim powerRow() As Long, powerCol() As Long
Dim powerType() As String, powerActive() As Boolean
Dim powerSpawnTime() As Double
Public activePowerUp As String
Public activePowerUpEndTime As Date

' Orange coin
Dim orangeRow As Long, orangeCol As Long, orangeValue As Long, orangeActive As Boolean
Public nextOrangeExpireTime As Date
Const ORANGE_LIFETIME As Double = 15

' Timers
Public nextUpdateTime As Date
Public nextPowerSpawnTime As Date

' Player
Public playerName As String

' Legend
Const LEGEND_START_ROW As Long = 14
Const LEGEND_START_COL As Long = 30

' Constants
Const MIN_INTERVAL As Double = 0.2
Const POWER_LIFETIME As Double = 15
Const POWER_SPAWN_INTERVAL As Double = 20

'running total of apples eaten for scoreboard
Public totalApplesEaten As Long

'running total of coins (powerups) eaten for scoreboard
Public totalCoins As Long

'======================
' CHASING SNAKE (LEVEL 20+)
'======================
Dim chaserR() As Integer, chaserC() As Integer
Dim chaserLength As Long
Dim chaserActive As Boolean
Dim chaserColor As Long
Public chaserInterval As Double
Public nextChaserUpdateTime As Date
'======================
' INITIALIZATION
'======================
Sub StartGame()
    Randomize
    running = True
    paused = False
    game = True
    
    BindKeys
    DrawLevelLegend
    
    ' --- Cheat Code Handler ---
    Dim cheatVal As String
    cheatVal = UCase(Trim(Range("AE25").Value))
    
    ' Reset cheat-dependent flags
    wallPassActive = False
    
    Select Case True
        Case cheatVal = "T"
            level = 5
            MsgBox "Cheat Activated! Starting at Level 5 (TRIS)", vbInformation, "Cheat Code"
        Case cheatVal = "S"
            wallPassActive = True
            MsgBox "Cheat Activated! WallPass is ON", vbInformation, "Cheat Code"
        Case IsNumeric(cheatVal)
            level = CLng(cheatVal)
            If level < 1 Then level = 1
            'If level > 20 Then level = 20
            MsgBox "Cheat Activated! Starting at Level " & level, vbInformation, "Cheat Code"
        Case Else
            level = 1
    End Select
    
     ' Update level display
        Range("B30").Value = "Level: " & level
        Range("AD23").Value = ""
    
    ' Clear the cheat input
    Range("AE25").Value = ""
    
    ' --- Reset apples / stats ---
    apples = 0: extraCoins = 0: applesEaten = 0
    totalCoins = 0
    totalApplesEaten = 0
    
    applesForNextLevel = 5
    moveInterval = 1
    baseSpeed = moveInterval
    activePowerUp = ""
    nextPowerSpawnTime = Now + TimeSerial(0, 0, POWER_SPAWN_INTERVAL)
    SetActivePowerUpDisplay
    
    ' --- Colours ---
    boardColor = RGB(255, 246, 211)
    snakeBodyColor = RGB(249, 168, 117)
    snakeheadColor = RGB(235, 107, 111)
    appleColor = RGB(0, 112, 192)
    wallColor = RGB(64, 64, 64)
    
    ' --- Initialize arrays ---
    ReDim r(0 To 2)
    ReDim c(0 To 2)
    
    Const MAX_POWERUPS As Long = 5
    ReDim powerRow(0 To MAX_POWERUPS - 1)
    ReDim powerCol(0 To MAX_POWERUPS - 1)
    ReDim powerType(0 To MAX_POWERUPS - 1)
    ReDim powerActive(0 To MAX_POWERUPS - 1)
    ReDim powerSpawnTime(0 To MAX_POWERUPS - 1)
    Dim i As Long
    For i = 0 To MAX_POWERUPS - 1
        powerActive(i) = False
    Next i
    
    ' --- Initialize board ---
    Set rng = Range("B2:Z27")
    rng.Clear
    rng.Interior.Color = boardColor
    
    ' --- Player input ---
    playerName = InputBox("Enter your name:", "Player Setup", "Snek")
    If Trim(playerName) = "" Then playerName = "Snek"
    
    ' --- Snake setup (bottom center) ---
    r(0) = 20: r(1) = 21: r(2) = 22
    c(0) = 14: c(1) = 14: c(2) = 14
    rinc = -1: cinc = 0
    
    ' --- Spawn initial items ---
    SpawnApple
    SpawnOrangeCoin
    
    ' --- Draw visuals ---
    ShowSnake
    DrawLegend
    SetScore
    
    ' --- Special levels ---
    If level = 5 Then StartSpecialLevel
    
    'walls between 11 and 12 (light)
    If level >= 11 And level <= 12 Then SpawnWalls 5 + (level * 2)
    
    'this activates the chaser at levels greater than 20
    If level >= 20 Then
        chaserActive = True
        chaserColor = RGB(0, 180, 180)
        ReDim chaserR(0 To 2)
        ReDim chaserC(0 To 2)
        chaserR(0) = 2: chaserC(0) = 26
        chaserR(1) = 3: chaserC(1) = 26
        chaserR(2) = 4: chaserC(2) = 26
        
        ' Determine initial chaser speed (change this if you want the chaser to go faster)
        If level >= 30 Then
            chaserInterval = moveInterval * 0.8
        Else
            chaserInterval = moveInterval
        End If
    End If
    
    ' --- Ensure safe moveInterval ---
    If moveInterval < MIN_INTERVAL Then moveInterval = MIN_INTERVAL
    
    ' --- Schedule game loop ---
    nextUpdateTime = Now + TimeSerial(0, 0, moveInterval)
    Application.OnTime nextUpdateTime, "Update"
    
    ' --- Schedule chaser loop if needed ---
    If chaserActive Then
        nextChaserUpdateTime = Now + TimeSerial(0, 0, chaserInterval)
        Application.OnTime nextChaserUpdateTime, "UpdateChaser"
    End If
End Sub
Sub FlashTrap(rw As Long, cl As Long)
    
    Cells(rw, cl).Interior.Color = RGB(255, 150, 150)
    
    ' Set up flash timing
    trapFlashRow = rw
    trapFlashCol = cl
    trapFlashEnd = Timer + 10
    trapFlashing = True
End Sub
Sub ResetGame()
    running = False
    paused = False
    game = False
    
    ' Cancel timers
    On Error Resume Next
    If nextUpdateTime <> 0 Then Application.OnTime nextUpdateTime, "Update", Schedule:=False
    If nextPowerSpawnTime <> 0 Then Application.OnTime nextPowerSpawnTime, "SpawnPowerUp", Schedule:=False
    If nextChaserUpdateTime <> 0 Then Application.OnTime nextChaserUpdateTime, "UpdateChaser", Schedule:=False

    On Error GoTo 0
    
    ' Clear board
    rng.Clear
    rng.Interior.Color = boardColor
    
    ' Reset variables
    apples = 0: extraCoins = 0: applesEaten = 0
    orangeActive = False
    wallCount = 0
    trapCount = 0
    totalApplesEaten = 0
    totalCoins = 0
    
    ReDim r(0 To 2)
    ReDim c(0 To 2)
    
    'stop chaser
    chaserActive = False
    
    SetScore
    Range("AD21").Value = "Active Power-Up: None"
    Range("AD22").Value = ""
    Range("AD23").Value = ""
    Range("B30").Value = "Level:"
    
    Unbindkeys
End Sub
Sub TogglePause()
    If Not game Then Exit Sub
    paused = Not paused
    If Not paused Then Update
End Sub
'======================
' SNAKE MOVEMENT
'======================
Sub MoveSnake()
    If rinc = 0 And cinc = 0 Then Exit Sub
    Dim tailIdx As Long, i As Long
    tailIdx = UBound(r)
    
    ' Clear tail
    Cells(r(tailIdx), c(tailIdx)).Interior.Color = boardColor
    Cells(r(tailIdx), c(tailIdx)).Value = ""
    
    ' Move body
    For i = tailIdx To 1 Step -1
        r(i) = r(i - 1)
        c(i) = c(i - 1)
    Next i
    
    ' Move head
    r(0) = r(0) + rinc
    c(0) = c(0) + cinc
    
    ' Border collision
    If r(0) < 2 Or r(0) > 27 Or c(0) < 2 Or c(0) > 26 Then
        If wallPassActive Then
            If r(0) < 2 Then r(0) = 27
            If r(0) > 27 Then r(0) = 2
            If c(0) < 2 Then c(0) = 26
            If c(0) > 26 Then c(0) = 2
        Else
            SnakeGameOverMsg "You hit the border."
            Exit Sub
        End If
    End If
    
    ' Self collision
    For i = 1 To UBound(r)
        If r(0) = r(i) And c(0) = c(i) Then
            SnakeGameOverMsg "You ran into yourself."
            Exit Sub
        End If
    Next i
    
    ' Wall collision
    For i = 0 To wallCount - 1
        If wallRows(i) = r(0) And wallCols(i) = c(0) Then
            If Not wallPassActive Then
                SnakeGameOverMsg "You hit a wall."
                Exit Sub
            End If
        End If
    Next i
    
    ' Trap collision (if traps exist)
If trapCount > 0 Then
    For i = 0 To trapCount - 1
        If r(0) = trapRows(i) And c(0) = trapCols(i) Then
            If applesEaten > 0 Then applesEaten = applesEaten - 1
            FlashTrap trapRows(i), trapCols(i)
        End If
    Next i
    Range("AD23") = "You hit a trap! Removed 1 apple"
End If

    ' --- Apple collision ---
    If Cells(r(0), c(0)).Interior.Color = appleColor Then
    apples = apples + 1
    totalApplesEaten = totalApplesEaten + 1
    applesEaten = applesEaten + 1
    ExtendSnake
    SpawnApple
    SetScore
    UpdateLiveScoreboard
    
    ' Check for level up
    If applesEaten >= applesForNextLevel Then
        LevelUp
    End If
    End If

    ' --- Display apples collected ---
    Range("AD22").Value = "Apples collected: " & applesEaten & " / " & applesForNextLevel
    Range("AD22").Interior.Color = boardColor
    Range("AD22").Font.Color = RGB(0, 0, 0)
    
    ' Orange collision
    If orangeActive Then
        If r(0) = orangeRow And c(0) = orangeCol Then
            apples = apples + orangeValue
            extraCoins = extraCoins + 1
            totalCoins = totalCoins + 1
            UpdateLiveScoreboard
            orangeActive = False
            Cells(orangeRow, orangeCol).Interior.Color = boardColor
            Cells(orangeRow, orangeCol).Value = ""
            SetScore
        End If
    End If
    
    ' Power-up collision
    For i = LBound(powerRow) To UBound(powerRow)
        If powerActive(i) Then
            If r(0) = powerRow(i) And c(0) = powerCol(i) Then
            
            totalCoins = totalCoins + 1
            
                ActivatePowerUp i
                Exit For
            End If
        End If
    Next i
    
    'updatescoreboard immediatly
    UpdateLiveScoreboard
    
    ' Draw snake
    ShowSnake
End Sub
Sub ExtendSnake()
    ReDim Preserve r(0 To UBound(r) + 1)
    ReDim Preserve c(0 To UBound(c) + 1)
    r(UBound(r)) = r(UBound(r) - 1)
    c(UBound(c)) = c(UBound(c) - 1)
End Sub
'======================
' LEVEL PROGRESSION
'======================
Sub LevelUp()
    level = level + 1
    
    If level > 31 Then level = 30 'stops game at end level instead of 30 for scoreboard
    
    ' Determine apples required for this level (change these settings if you want things to be harder)
    Select Case level
        Case 1 To 4
            applesForNextLevel = 5
            applesEaten = 0
        Case 5
            applesForNextLevel = 10
            applesEaten = 0
            StartSpecialLevel
        Case 6 To 10
            ClearWalls
            applesForNextLevel = 5
            applesEaten = 0
        Case 11 To 12
            ClearWalls
            applesForNextLevel = 5
            applesEaten = 0
            SpawnWalls 5 + (level * 2)
        Case 13 To 14
            ClearWalls
            applesForNextLevel = 5
            applesEaten = 0
            SpawnWalls 5 + (level * 4)
        Case 15 To 16
            ClearWalls
            applesForNextLevel = 5
            applesEaten = 0
            SpawnWalls 5 + (level * 1)
        Case 17 To 17
            ClearWalls
            applesForNextLevel = 5
            applesEaten = 0
            MsgBox "Lets drop in some traps for level 17 and 18, which remove 1 apple", vbInformation, "Traps Mode"
            SpawnTraps 10
        Case 18 To 18
            ClearWalls
            applesForNextLevel = 5
            applesEaten = 0
            SpawnTraps 10
        Case 19 To 19
            ClearWalls
            applesForNextLevel = 10
            applesEaten = 0
            StartSpecialLevel
        Case 20 To 22
            applesForNextLevel = 5
            applesEaten = 0
            ClearWalls
            ClearChaser
            SpawnApple
            chaserActive = True
            chaserColor = RGB(0, 180, 180)
        ReDim chaserR(0 To 2)
        ReDim chaserC(0 To 2)
        chaserR(0) = 2: chaserC(0) = 26
        chaserR(1) = 3: chaserC(1) = 26
        chaserR(2) = 4: chaserC(2) = 26
        Case 23 To 25
            applesForNextLevel = 5
            applesEaten = 0
        Case 26 To 28
            applesForNextLevel = 5
            applesEaten = 0
            SpawnWalls 5 + (level * 1)
        Case 29 To 29
            applesEaten = 0
            applesForNextLevel = 10
            SpawnWalls 5 + (level * 3)
            ClearChaser
            SpawnApple
        Case 30
        ' FINAL LEVEL 30 - The Ultimate Challenge
            SpawnTraps 5
            applesForNextLevel = 25
            applesEaten = 0
            ClearWalls
            ClearTraps
            SpawnWalls 5 + (level * 1)
            ClearChaser
            SpawnApple
            chaserActive = True
            chaserColor = RGB(0, 180, 180)
            ReDim chaserR(0 To 2)
        ReDim chaserC(0 To 2)
            chaserR(0) = 2: chaserC(0) = 26
            chaserR(1) = 3: chaserC(1) = 26
            chaserR(2) = 4: chaserC(2) = 26
            chaserInterval = moveInterval * 0.9
            MsgBox "LEVEL 30: The Final Challenge!" & vbCrLf & _
           "Survive the Chaser King and collect 25 apples to win!", vbExclamation, "FINAL LEVEL"
            level = 30
        Case Else
            applesForNextLevel = 5
            applesEaten = 0
            SpawnWalls 5 + (level * 2)
            Range("AD23").Value = ""
            
    End Select
    Range("AD23").Value = ""
    
    ' Update display
    Range("AD22").Value = "Apples collected: 0 / " & applesForNextLevel
    UpdateLiveScoreboard
    
    ' Speed up snake
    moveInterval = moveInterval * 0.9
    If moveInterval < MIN_INTERVAL Then moveInterval = MIN_INTERVAL

    ' Power-ups
    If level >= 15 Then
        SpawnPowerUp
    Else
        SpawnPowerUp
    End If
    
    ' Traps
    If level >= 17 And level <= 18 Then
    Range("AD23") = "Traps are active! Be Careful"
    SpawnTraps 10
    Else
        ClearTraps
    End If
    ' Power-ups
    If level >= 20 Then
        SpawnPowerUp
        SpawnPowerUp
    Else
        SpawnPowerUp
    End If
    
End Sub
'======================
' SPECIAL LEVEL (TRIS 2025)
'======================
Sub StartSpecialLevel()
        
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    rng.Clear
    rng.Interior.Color = boardColor
    
    ' Reset snake
    r(0) = 22: c(0) = 14
    r(1) = 23: c(1) = 14
    r(2) = 24: c(2) = 14
    rinc = -1: cinc = 0
    ShowSnake
    
    wallCount = 0
    Dim tempWalls() As Variant, idx As Long
    idx = 0
    ReDim tempWalls(0 To 400)
    
    Dim baseRow As Long, baseCol As Long
    Dim rOffset As Long, cOffset As Long
    Dim i As Long
    
    ' ====== TRIS (Top) ======
    baseRow = 4
    
    ' T
    baseCol = 3
    For cOffset = 0 To 4
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
    Next cOffset
    For rOffset = 1 To 6
        tempWalls(idx) = Array(baseRow + rOffset, baseCol + 2): idx = idx + 1
    Next rOffset
    
    ' R
    baseCol = 10
    For rOffset = 0 To 6
        tempWalls(idx) = Array(baseRow + rOffset, baseCol): idx = idx + 1
    Next rOffset
    For cOffset = 1 To 3
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 3, baseCol + 1): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 3, baseCol + 2): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 1, baseCol + 3): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 2, baseCol + 3): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 4, baseCol + 2): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 5, baseCol + 3): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 6, baseCol + 4): idx = idx + 1
    
    ' I (shorter)
    baseCol = 17
    For cOffset = 0 To 2
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
    Next cOffset
    For rOffset = 1 To 5
        tempWalls(idx) = Array(baseRow + rOffset, baseCol + 1): idx = idx + 1
    Next rOffset
    For cOffset = 0 To 2
        tempWalls(idx) = Array(baseRow + 6, baseCol + cOffset): idx = idx + 1
    Next cOffset
    
    ' S
    baseCol = 22
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 1, baseCol): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 2, baseCol): idx = idx + 1
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow + 3, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 4, baseCol + 3): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 5, baseCol + 3): idx = idx + 1
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow + 6, baseCol + cOffset): idx = idx + 1
    Next cOffset
    
    ' ====== 2025 (Bottom) ======
    baseRow = 14
    
    ' 2
    baseCol = 3
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 1, baseCol + 3): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 2, baseCol + 2): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 3, baseCol + 1): idx = idx + 1
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow + 4, baseCol + cOffset): idx = idx + 1
    Next cOffset
    
    ' 0
    baseCol = 9
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
        tempWalls(idx) = Array(baseRow + 4, baseCol + cOffset): idx = idx + 1
    Next cOffset
    For rOffset = 1 To 3
        tempWalls(idx) = Array(baseRow + rOffset, baseCol): idx = idx + 1
        tempWalls(idx) = Array(baseRow + rOffset, baseCol + 3): idx = idx + 1
    Next rOffset
    
    ' 2
    baseCol = 16
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 1, baseCol + 3): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 2, baseCol + 2): idx = idx + 1
    tempWalls(idx) = Array(baseRow + 3, baseCol + 1): idx = idx + 1
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow + 4, baseCol + cOffset): idx = idx + 1
    Next cOffset
    
    ' 5
    baseCol = 22
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow, baseCol + cOffset): idx = idx + 1
        tempWalls(idx) = Array(baseRow + 4, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 1, baseCol): idx = idx + 1
    For cOffset = 0 To 3
        tempWalls(idx) = Array(baseRow + 2, baseCol + cOffset): idx = idx + 1
    Next cOffset
    tempWalls(idx) = Array(baseRow + 3, baseCol + 3): idx = idx + 1
    
    ' ====== Place walls ======
    ReDim Preserve tempWalls(0 To idx - 1)
    wallCount = UBound(tempWalls) + 1
    ReDim wallRows(0 To wallCount - 1)
    ReDim wallCols(0 To wallCount - 1)
    For i = 0 To wallCount - 1
        wallRows(i) = tempWalls(i)(0)
        wallCols(i) = tempWalls(i)(1)
        ws.Cells(wallRows(i), wallCols(i)).Interior.Color = wallColor
    Next i
    
    ' Add gameplay items
    SpawnApple
    SpawnApple
    SpawnPowerUp
    SpawnOrangeCoin
    
    moveInterval = baseSpeed * 0.6
    SetScore
    UpdateLiveScoreboard
End Sub
'======================
' BOARD / SNAKE DRAWING
'======================
Sub ShowSnake()
    Dim i As Long
    For i = 1 To UBound(r)
        Cells(r(i), c(i)).Interior.Color = snakeBodyColor
    Next i
    Cells(r(0), c(0)).Interior.Color = snakeheadColor
End Sub

Sub DrawLegend()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim startRow As Long, startCol As Long
    Dim boardColor As Long, headerColor As Long

    startRow = 14
    startCol = 33

    ' Pull colors dynamically from sheet
    boardColor = ws.Range("B2").Interior.Color
    headerColor = ws.Range("AD2").Interior.Color

    ' === Clear existing legend area ===
    ws.Range(ws.Cells(startRow - 1, startCol), ws.Cells(startRow + 5, startCol + 1)).ClearContents
    ws.Range(ws.Cells(startRow - 1, startCol), ws.Cells(startRow + 5, startCol + 1)).Interior.ColorIndex = xlNone

    ' === Header row ("Type" / "Color") ===
    ws.Cells(startRow - 1, startCol).Value = "Type"
    ws.Cells(startRow - 1, startCol + 1).Value = "Color"

    With ws.Range(ws.Cells(startRow - 1, startCol), ws.Cells(startRow - 1, startCol + 1))
        .Interior.Color = headerColor
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' === Labels ===
    ws.Cells(startRow, startCol).Value = "Snake Head"
    ws.Cells(startRow + 1, startCol).Value = "Snake Body"
    ws.Cells(startRow + 2, startCol).Value = "Apple"
    ws.Cells(startRow + 3, startCol).Value = "Special"
    ws.Cells(startRow + 4, startCol).Value = "Traps"
    ws.Cells(startRow + 5, startCol).Value = "Enemy"

    ' === Fill left column background to match board ===
    ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 5, startCol)).Interior.Color = boardColor

    ' === Color squares ===
    ws.Cells(startRow, startCol + 1).Interior.Color = snakeheadColor        ' Snake Head
    ws.Cells(startRow + 1, startCol + 1).Interior.Color = snakeBodyColor    ' Snake Body
    ws.Cells(startRow + 2, startCol + 1).Interior.Color = appleColor    ' Apple
    ws.Cells(startRow + 3, startCol + 1).Interior.Color = RGB(255, 255, 0)  ' Special
    ws.Cells(startRow + 4, startCol + 1).Interior.Color = RGB(169, 251, 104) ' Traps
    ws.Cells(startRow + 5, startCol + 1).Interior.Color = RGB(0, 180, 180) 'Enemy Snek

    ' === Borders & alignment ===
    With ws.Range(ws.Cells(startRow - 1, startCol), ws.Cells(startRow + 4, startCol + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub
Sub SetScore()
    Range("B29").Value = "Score: " & apples & "   " & "Extra Coins: " & extraCoins
End Sub
'======================
' APPLES / ORANGE
'======================
Sub SpawnApple()
    Dim rowA As Long, colA As Long
    Do
        rowA = Int((27 - 2 + 1) * Rnd + 2)
        colA = Int((26 - 2 + 1) * Rnd + 2)
    Loop While Cells(rowA, colA).Interior.Color <> boardColor
    Cells(rowA, colA).Interior.Color = appleColor
    Cells(rowA, colA).Value = ""
End Sub

Sub SpawnOrangeCoin()
    orangeActive = True
    orangeRow = Int((27 - 2 + 1) * Rnd + 2)
    orangeCol = Int((26 - 2 + 1) * Rnd + 2)
    orangeValue = 5
    Cells(orangeRow, orangeCol).Interior.Color = RGB(255, 165, 0)
    nextOrangeExpireTime = Now + TimeSerial(0, 0, ORANGE_LIFETIME)
End Sub

Sub CheckOrangeExpire()
    If orangeActive Then
        If Now >= nextOrangeExpireTime Then
            Cells(orangeRow, orangeCol).Interior.Color = boardColor
            Cells(orangeRow, orangeCol).Value = ""
            orangeActive = False
        End If
    End If
End Sub

'======================
' WALLS
'======================
Sub SpawnWalls(count As Long)
    Dim i As Long, rw As Long, cl As Long
    ClearWalls
    wallCount = count
    ReDim wallRows(0 To wallCount - 1)
    ReDim wallCols(0 To wallCount - 1)
    For i = 0 To wallCount - 1
        Do
            rw = Int((27 - 2 + 1) * Rnd + 2)
            cl = Int((26 - 2 + 1) * Rnd + 2)
        Loop While Cells(rw, cl).Interior.Color <> boardColor
        wallRows(i) = rw
        wallCols(i) = cl
        Cells(rw, cl).Interior.Color = wallColor
    Next i
End Sub

Sub ClearWalls()
    Dim i As Long
    If wallCount > 0 Then
        For i = 0 To wallCount - 1
            Cells(wallRows(i), wallCols(i)).Interior.Color = boardColor
        Next i
    End If
    wallCount = 0
End Sub

'======================
' MOVING WALLS (Level 12+)
'======================
Sub MoveWalls()
    Dim i As Long, dir As Long, j As Long
    If wallCount = 0 Then Exit Sub

    For i = 0 To wallCount - 1
        ' Clear old wall
        Cells(wallRows(i), wallCols(i)).Interior.Color = boardColor

        ' Random horizontal direction: -1 or 1
        dir = Choose(Int(2 * Rnd) + 1, -1, 1)
        wallCols(i) = wallCols(i) + dir

        ' Keep wall in bounds
        If wallCols(i) < 2 Then wallCols(i) = 2
        If wallCols(i) > 26 Then wallCols(i) = 26

        ' Draw new wall
        Cells(wallRows(i), wallCols(i)).Interior.Color = wallColor

        ' --- Collision check: did wall move into the snake? ---
        For j = 0 To UBound(r)
            If r(j) = wallRows(i) And c(j) = wallCols(i) Then
                SnakeGameOverMsg "A wall moved into you!"
                Exit Sub
            End If
        Next j
    Next i
End Sub
'======================
' TRAP CELLS
'======================
Sub SpawnTraps(count As Long)
    Dim i As Long, rw As Long, cl As Long
    Dim startTime As Double
    
    ' Clear any existing traps
    ClearTraps
    
    trapCount = count
    ReDim trapRows(0 To trapCount - 1)
    ReDim trapCols(0 To trapCount - 1)
    
    For i = 0 To trapCount - 1
        ' Pick a random cell that matches the board color
        Do
            rw = Int((27 - 2 + 1) * Rnd + 2)
            cl = Int((26 - 2 + 1) * Rnd + 2)
        Loop While Cells(rw, cl).Interior.Color <> boardColor
        
        trapRows(i) = rw
        trapCols(i) = cl
        
        ' Flash trap briefly
        Cells(rw, cl).Interior.Color = RGB(169, 251, 104)
        
        ' Use Timer loop for sub-second delay
        startTime = Timer
        Do While Timer < startTime + 0.2
            DoEvents
        Loop
        
        ' Revert to normal board color
        Cells(rw, cl).Interior.Color = boardColor
    Next i
End Sub
Sub UpdateTrapFlash()
    If trapFlashing Then
        If Timer >= trapFlashEnd Then
            ' Revert cell back to normal board color
            Cells(trapFlashRow, trapFlashCol).Interior.Color = boardColor
            trapFlashing = False
        End If
    End If
End Sub
Sub CheckTrapCollision()
    Dim i As Long
    If trapCount > 0 Then
        For i = 0 To trapCount - 1
            If r(0) = trapRows(i) And c(0) = trapCols(i) Then
                If applesEaten > 0 Then applesEaten = applesEaten - 1
                FlashTrap trapRows(i), trapCols(i)
            End If
        Next i
    End If
End Sub
Sub ClearTraps()
    Dim i As Long
    If trapCount > 0 Then
        For i = 0 To trapCount - 1
            Cells(trapRows(i), trapCols(i)).Interior.Color = boardColor
        Next i
    End If
    trapCount = 0
End Sub
Sub SuperShrink()
    Const ORIGINAL_SIZE As Long = 3 'change this if you want the snake to shrink to a different size
    Dim oldLength As Long
    Dim i As Long

    oldLength = UBound(r) + 1

    If oldLength > ORIGINAL_SIZE Then

        For i = ORIGINAL_SIZE To oldLength - 1
            Cells(r(i), c(i)).Interior.Color = boardColor
            Cells(r(i), c(i)).Value = ""
        Next i

        ReDim Preserve r(0 To ORIGINAL_SIZE - 1)
        ReDim Preserve c(0 To ORIGINAL_SIZE - 1)

        ShowSnake
    End If
End Sub
'======================
' POWER-UPS
'======================
Sub SpawnPowerUp()
    Dim idx As Long
    For idx = LBound(powerRow) To UBound(powerRow)
        If Not powerActive(idx) Then
            powerActive(idx) = True
            powerRow(idx) = Int((27 - 2 + 1) * Rnd + 2)
            powerCol(idx) = Int((26 - 2 + 1) * Rnd + 2)
            powerType(idx) = Choose(Int(6 * Rnd) + 1, "Speed", "Slow", "WallPass", "Shrink", "Bonus", "SuperShrink")
            powerSpawnTime(idx) = Now
            
            Select Case powerType(idx)
        Case "Speed": Cells(powerRow(idx), powerCol(idx)).Interior.Color = RGB(255, 0, 0)
        Case "Slow": Cells(powerRow(idx), powerCol(idx)).Interior.Color = RGB(0, 0, 255)
        Case "WallPass": Cells(powerRow(idx), powerCol(idx)).Interior.Color = RGB(0, 128, 0)
        Case "Shrink": Cells(powerRow(idx), powerCol(idx)).Interior.Color = RGB(255, 128, 0)
        Case "Bonus": Cells(powerRow(idx), powerCol(idx)).Interior.Color = RGB(255, 215, 0)
        Case "SuperShrink": Cells(powerRow(idx), powerCol(idx)).Interior.Color = RGB(128, 0, 128)
        End Select
            Exit For
        End If
    Next idx
    nextPowerSpawnTime = Now + TimeSerial(0, 0, POWER_SPAWN_INTERVAL)
End Sub
Sub SetActivePowerUpDisplay()
    Dim remaining As Long
    Dim displayText As String
    
    If activePowerUp <> "" Then
        remaining = Round((activePowerUpEndTime - Now) * 86400, 0)
        If remaining < 0 Then remaining = 0
        displayText = "Active Power-Up: " & activePowerUp & " (" & remaining & "s left)"
        
        Range("AD21").Value = displayText
        DoEvents
                
        With Range("AD21").Font
            Select Case activePowerUp
                Case "Speed": .Color = RGB(255, 0, 0)
                Case "Slow": .Color = RGB(0, 0, 255)
                Case "WallPass": .Color = RGB(0, 128, 0)
                Case "Shrink": .Color = RGB(255, 128, 0)
                Case "Bonus": .Color = RGB(255, 215, 0)
                Case Else: .Color = vbBlack
            End Select
        End With
    Else
        Range("AD21").Value = "Active Power-Up: None"
        Range("AD21").Font.Color = vbBlack
    End If
End Sub
Sub ActivatePowerUp(idx As Long)
    activePowerUp = powerType(idx)
    activePowerUpEndTime = Now + TimeSerial(0, 0, POWER_LIFETIME)
    powerActive(idx) = False
    Cells(powerRow(idx), powerCol(idx)).Interior.Color = boardColor
    Cells(powerRow(idx), powerCol(idx)).Value = ""
    
    Select Case activePowerUp
        Case "Speed"
            moveInterval = baseSpeed * 0.75
            If moveInterval < MIN_INTERVAL Then moveInterval = MIN_INTERVAL
        Case "Slow"
            moveInterval = baseSpeed * 1.5
        Case "WallPass"
            wallPassActive = True
        Case "Shrink"
            ShrinkSnake
        Case "Bonus"
            apples = apples + 3
            SetScore
        Case "SuperShrink"
        SuperShrink
    End Select
    
    SetActivePowerUpDisplay
End Sub

Sub ShrinkSnake()
    Const SHRINK_AMOUNT As Long = 3   ' change this if you want to remove more segments
    Dim oldLength As Long
    Dim newLength As Long
    Dim i As Long

    oldLength = UBound(r) + 1
    newLength = oldLength - SHRINK_AMOUNT

    If newLength < 3 Then newLength = 3

    For i = newLength To oldLength - 1
        Cells(r(i), c(i)).Interior.Color = boardColor
        Cells(r(i), c(i)).Value = ""
    Next i

    ReDim Preserve r(0 To newLength - 1)
    ReDim Preserve c(0 To newLength - 1)

    ShowSnake
End Sub



Sub CheckPowerExpire()
    If activePowerUp <> "" Then
        If Now >= activePowerUpEndTime Then
            Select Case activePowerUp
                Case "Speed", "Slow": moveInterval = baseSpeed
                Case "WallPass": wallPassActive = False
            End Select
            activePowerUp = ""
            SetActivePowerUpDisplay
        End If
    End If
End Sub
'======================
' GAME LOOP
'======================
Sub Update()
    If Not running Or paused Then Exit Sub

    ' --- Move walls first (level 12+) ---
    If level >= 15 And level <= 16 Then MoveWalls

    If level = 31 Then
    rng.Clear
    rng.Interior.Color = RGB(255, 255, 255)
    MsgBox "You've reached the end — TRIS salutes you!", vbInformation, "The End"
    Range("B30").Value = "Level: " & " The End"
    Unbindkeys
    chaserActive = True
    ClearChaser
    DrawVictoryFace
    Range("ad22") = ""
    Exit Sub
    End If

    ' Handle trap flash
    If trapFlashing Then
    If Timer >= trapFlashEnd Then
        Cells(trapFlashRow, trapFlashCol).Interior.Color = boardColor
        trapFlashing = False
    End If
    End If

    MoveSnake
    
    ' --- Move chasing snake if active and level >= 20 ---
    If chaserActive And level >= 20 Then MoveChaser

    CheckOrangeExpire
    CheckPowerExpire
    SetActivePowerUpDisplay

    nextUpdateTime = Now + TimeSerial(0, 0, moveInterval)
    Application.OnTime nextUpdateTime, "Update"
    
    
    ' --- Power-up spawn ---
    If nextPowerSpawnTime = 0 Or Now >= nextPowerSpawnTime Then
        SpawnPowerUp
    End If

    ' Debug info (remove this if you don't need it anymore)
    Debug.Print "Update running at " & Now & " Level: " & level
End Sub
'======================
' GAME OVER
'======================
Sub SnakeGameOverMsg(reason As String)
    running = False
    paused = False
    MsgBox "Game Over: " & reason & vbCrLf & "Final Score: " & apples, vbExclamation, "Snake"
    ResetGame
    UpdateHighScore
End Sub
'======================
' SCOREBOARD
'======================
Sub UpdateScoreboard()
    Dim rowIdx As Long
    rowIdx = 3 ' row to write player data

    With ActiveSheet
        .Range("AD" & rowIdx).Value = playerName
        .Range("AE" & rowIdx).Value = apples

        ' Clamp level display so it never exceeds 30
        .Range("AF" & rowIdx).Value = IIf(level > 30, 30, level)

        .Range("AG" & rowIdx).Value = totalApplesEaten
        .Range("AH" & rowIdx).Value = totalCoins

        ' Maintain consistent styling
        .Range("AD" & rowIdx & ":AH" & rowIdx).Interior.Color = boardColor
        .Range("AD" & rowIdx & ":AH" & rowIdx).Font.Color = RGB(0, 0, 0)
    End With
    
    ' --- Correct any accidental "31" display (Remove this if you don't need it as I've clamped it elsewhere in levelup ---
    Dim cell As Range
    For Each cell In ActiveSheet.Range("AF3:AF12")
    If cell.Value = 31 Then cell.Value = 30
    Next cell
    
End Sub
Sub UpdateHighScore()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim firstRow As Long, lastRow As Long
    firstRow = 3
    lastRow = 12
    
    Dim i As Long, found As Boolean
    found = False
    
    ' Update existing
    For i = firstRow To lastRow
        If ws.Cells(i, "AD").Value = playerName Then
            If apples > ws.Cells(i, "AE").Value Then
                ws.Cells(i, "AE").Value = apples
                ws.Cells(i, "AF").Value = level
                ws.Cells(i, "AG").Value = applesEaten
                ws.Cells(i, "AH").Value = totalCoins
            End If
            found = True
            Exit For
        End If
    Next i
    
    ' Add or replace lowest
    If Not found Then
        Dim minScore As Long, minRow As Long
        minScore = ws.Cells(firstRow, "AE").Value
        minRow = firstRow
        For i = firstRow To lastRow
            If ws.Cells(i, "AE").Value < minScore Or ws.Cells(i, "AE").Value = "" Then
                minScore = IIf(ws.Cells(i, "AE").Value = "", -1, ws.Cells(i, "AE").Value)
                minRow = i
            End If
        Next i
        
        ws.Cells(minRow, "AD").Value = playerName
        ws.Cells(minRow, "AE").Value = apples
        ws.Cells(minRow, "AF").Value = level
        ws.Cells(minRow, "AG").Value = applesEaten
        ws.Cells(minRow, "AH").Value = totalCoins
    End If
    
    ' Sort
    ws.Range("AD2:AH12").Sort Key1:=ws.Range("AE3"), Order1:=xlDescending, Header:=xlYes
    
    ' Format & highlight
    ws.Range("AD3:AH12").Interior.Color = boardColor
    ws.Range("AD3:AH12").Font.Color = RGB(0, 0, 0)
    For i = firstRow To lastRow
        If ws.Cells(i, "AD").Value = playerName Then
            ws.Range("AD" & i & ":AH" & i).Interior.Color = RGB(255, 255, 150)
        End If
    Next i
End Sub
Sub UpdateLiveScoreboard()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim firstRow As Long, lastRow As Long
    firstRow = 3
    lastRow = 12
    
    Dim i As Long, found As Boolean
    found = False
    
    ' Update existing
    For i = firstRow To lastRow
        If ws.Cells(i, "AD").Value = playerName Then
            ws.Cells(i, "AE").Value = apples
            ws.Cells(i, "AF").Value = level
            ws.Cells(i, "AG").Value = totalApplesEaten
            ws.Cells(i, "AH").Value = totalCoins
            found = True
            Exit For
        End If
    Next i
    
    ' Add or replace lowest
    If Not found Then
        Dim minScore As Long, minRow As Long
        minScore = ws.Cells(firstRow, "AE").Value
        minRow = firstRow
        For i = firstRow To lastRow
            If ws.Cells(i, "AE").Value < minScore Or ws.Cells(i, "AE").Value = "" Then
                minScore = IIf(ws.Cells(i, "AE").Value = "", -1, ws.Cells(i, "AE").Value)
                minRow = i
            End If
        Next i
        
        ws.Cells(minRow, "AD").Value = playerName
        ws.Cells(minRow, "AE").Value = apples
        ws.Cells(minRow, "AF").Value = level
        ws.Cells(minRow, "AG").Value = applesEaten
        ws.Cells(minRow, "AH").Value = totalCoins
    End If
    
    ' Format & highlight
    ws.Range("AD3:AH12").Interior.Color = boardColor
    ws.Range("AD3:AH12").Font.Color = RGB(0, 0, 0)
    For i = firstRow To lastRow
        If ws.Cells(i, "AD").Value = playerName Then
            ws.Range("AD" & i & ":AH" & i).Interior.Color = RGB(255, 255, 150)
        End If
    Next i
    
    'remove level 31 (another catch all for level 31 bug, can probably remove this now)
    Dim cell As Range
    For Each cell In ActiveSheet.Range("AF3:AF12")
    If cell.Value = 31 Then cell.Value = 30
    Next cell
    
    ' Sort
    ws.Range("AD2:AH12").Sort Key1:=ws.Range("AE3"), Order1:=xlDescending, Header:=xlYes
    
    ' Update level display
    Range("B30").Value = "Level: " & level
End Sub
Sub ClearScores()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("AD3:AH12").ClearContents
    ws.Range("AD3:AH12").Interior.Color = boardColor
    Range("AD21").Value = "Active Power-Up: None"
    Range("AD21").Font.Color = RGB(0, 0, 0)
End Sub

Sub ShowChaser()
    If Not chaserActive Then Exit Sub
    Dim i As Long
    ' Clear old chaser positions first
    For i = 0 To UBound(chaserR)
        Cells(chaserR(i), chaserC(i)).Interior.Color = boardColor
    Next i
    ' Draw chaser body
    For i = 1 To UBound(chaserR)
        Cells(chaserR(i), chaserC(i)).Interior.Color = chaserColor
    Next i
    ' Draw chaser head darker
    Cells(chaserR(0), chaserC(0)).Interior.Color = RGB(80, 0, 80)
End Sub
Sub MoveChaser()
    If Not chaserActive Then Exit Sub
    
    Dim i As Long, tailIdx As Long
    tailIdx = UBound(chaserR)
    
    ' Clear tail
    Cells(chaserR(tailIdx), chaserC(tailIdx)).Interior.Color = boardColor
    
    ' Move body
    For i = tailIdx To 1 Step -1
        chaserR(i) = chaserR(i - 1)
        chaserC(i) = chaserC(i - 1)
    Next i
    
    ' Candidate moves: Up, Down, Left, Right
    Dim drOptions(1 To 4) As Long, dcOptions(1 To 4) As Long
    drOptions(1) = -1: dcOptions(1) = 0
    drOptions(2) = 1: dcOptions(2) = 0
    drOptions(3) = 0: dcOptions(3) = -1
    drOptions(4) = 0: dcOptions(4) = 1
    
    Dim validMoves() As Variant
    ReDim validMoves(1 To 4)
    Dim moveCount As Long: moveCount = 0
    
    Dim newR As Long, newC As Long, blocked As Boolean
    Dim j As Long
    
    ' Gather all valid moves
    For i = 1 To 4
        newR = chaserR(0) + drOptions(i)
        newC = chaserC(0) + dcOptions(i)
        
        ' Bounds check
        If newR < 2 Or newR > 27 Or newC < 2 Or newC > 26 Then GoTo NextMove
        
        ' Check walls
        blocked = False
        For j = 0 To wallCount - 1
            If newR = wallRows(j) And newC = wallCols(j) Then blocked = True
        Next j
        
        ' Check chaser body
        For j = 1 To UBound(chaserR)
            If newR = chaserR(j) And newC = chaserC(j) Then blocked = True
        Next j
        
        If Not blocked Then
            moveCount = moveCount + 1
            validMoves(moveCount) = Array(newR, newC)
        End If
        
NextMove:
    Next i
    
    ' If no valid moves, stay in place
    If moveCount = 0 Then
        chaserR(0) = chaserR(1)
        chaserC(0) = chaserC(1)
    Else
        ' Compute distances to target
        Dim bestDist As Double: bestDist = 1000
        Dim candidateMoves() As Variant
        Dim candidateCount As Long: candidateCount = 0
        Dim dist As Double
        Dim targetR As Long, targetC As Long
        
        ' Determine target: player's head or predicted future position
        If level >= 25 Then
            ' Aggressive mode: predict 1-2 steps ahead
            targetR = r(0) + rinc * 2
            targetC = c(0) + cinc * 2
        Else
            targetR = r(0)
            targetC = c(0)
        End If
        
        For i = 1 To moveCount
            dist = Abs(validMoves(i)(0) - targetR) + Abs(validMoves(i)(1) - targetC)
            ' Allow slight randomness
            If dist <= bestDist + Int(2 * Rnd) Then
                bestDist = dist
                candidateCount = candidateCount + 1
                ReDim Preserve candidateMoves(1 To candidateCount)
                candidateMoves(candidateCount) = validMoves(i)
            End If
        Next i
        
        ' Pick one move randomly among candidates
        Dim pick As Long
        pick = Int(candidateCount * Rnd) + 1
        chaserR(0) = candidateMoves(pick)(0)
        chaserC(0) = candidateMoves(pick)(1)
    End If
    
    ' Collision with player snake = game over
    For i = 0 To UBound(r)
        If chaserR(0) = r(i) And chaserC(0) = c(i) Then
            SnakeGameOverMsg "The chasing snake caught you!"
            Exit Sub
        End If
    Next i
    
    ShowChaser
End Sub

Sub UpdateChaser()
    If Not running Or paused Or Not chaserActive Then Exit Sub
    
    MoveChaser
    
    ' Schedule next chaser move
    nextChaserUpdateTime = Now + TimeSerial(0, 0, chaserInterval)
    Application.OnTime nextChaserUpdateTime, "UpdateChaser"
End Sub

Sub DrawLevelLegend()
    Dim ws As Worksheet
    Dim startRow As Long
    Dim headerColor As Long
    Dim boardColor As Long
    Dim boardRowHeight As Double
    Dim data As Variant
    Dim i As Long
    
    Set ws = ActiveSheet
    startRow = 2
    
    ' Colors
    headerColor = ws.Range("AD2").Interior.Color
    boardColor = ws.Range("B2").Interior.Color
    
    ' Get board row height dynamically
    boardRowHeight = ws.Rows(2).RowHeight
    
    ' Unmerge cells to avoid auto-resizing
    ws.Range("AJ1:AK11").UnMerge
    
    ' Clear previous contents
    ws.Range("AJ2:AK11").ClearContents
    
    ' Disable wrap text BEFORE writing values
    ws.Range("AJ2:AK11").WrapText = False
    
    ' --- Headers ---
    ws.Cells(startRow, 36).Value = "Level"
    ws.Cells(startRow, 37).Value = "Type"
    ws.Cells(startRow, 36).Interior.Color = headerColor
    ws.Cells(startRow, 37).Interior.Color = headerColor
    ws.Cells(startRow, 36).Font.Bold = True
    ws.Cells(startRow, 37).Font.Bold = True
    ws.Rows(startRow).HorizontalAlignment = xlCenter
    ws.Rows(startRow).VerticalAlignment = xlCenter
    
    ' --- Data ---

    data = Array( _
    Array("1-4", "Normal"), _
    Array("5", "Special Level"), _
    Array("6-10", "Normal"), _
    Array("11-12", "Walls (Light)"), _
    Array("13-14", "Walls (Medium)"), _
    Array("15-16", "Moving Walls"), _
    Array("17-18", "10 Hidden Traps (-1 Apple)"), _
    Array("19", "Special Level"), _
    Array("20-22", "Snek Chaser Mode + More PowerUps"), _
    Array("23-25", "Faster Chaser Snek"), _
    Array("26-28", "Faster Chaser Snek + Walls"), _
    Array("29", "More Walls + More Apples"), _
    Array("30", "Walls + Traps + Faster Chase Snek") _
)
    
    For i = LBound(data) To UBound(data)
        ws.Cells(startRow + i + 1, 36).Value = data(i)(0)
        ws.Cells(startRow + i + 1, 37).Value = data(i)(1)
        ws.Cells(startRow + i + 1, 36).Interior.Color = boardColor
        ws.Cells(startRow + i + 1, 37).Interior.Color = boardColor
        ws.Cells(startRow + i + 1, 36).HorizontalAlignment = xlCenter
        ws.Cells(startRow + i + 1, 37).HorizontalAlignment = xlCenter
        ws.Cells(startRow + i + 1, 36).VerticalAlignment = xlCenter
        ws.Cells(startRow + i + 1, 37).VerticalAlignment = xlCenter
    Next i
    
    ' --- Borders & Font ---
    With ws.Range("AJ" & startRow & ":AK" & (startRow + UBound(data) + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
        .Font.Size = 10
    End With
    
    ' --- Column widths ---
    ws.Columns("AJ").ColumnWidth = 8
    ws.Columns("AK").ColumnWidth = 29.86
    
    ' --- Match board row height ---
    Dim r As Long
    For r = startRow To startRow + UBound(data) + 1
        ws.Rows(r).RowHeight = boardRowHeight
    Next r
    
' ==================================================================
' POWER-UP LEGEND
' ==================================================================
Dim pData As Variant
Dim pStartRow As Long
pStartRow = 14

' === Add Header Row (same style as scoreboard headers) ===
With ws.Range("AD13:AF13")
    .Value = Array("Power Up", "Description", "Colour")
    .Font.Bold = True
    .Font.Color = RGB(0, 0, 0)
    .Interior.Color = headerColor
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(180, 180, 180)
    .RowHeight = boardRowHeight
End With

' === Define Power-Up Data ===
pData = Array( _
    Array("Speed", "Increases movement speed", RGB(255, 0, 0)), _
    Array("Slow", "Slows down movement", RGB(0, 0, 255)), _
    Array("WallPass", "Pass through walls safely", RGB(0, 128, 0)), _
    Array("Shrink", "Shrinks the snake a small amount", RGB(255, 128, 0)), _
    Array("Bonus", "Gives +3 apples instantly", RGB(255, 215, 0)), _
    Array("SuperShrink", "Resets snake to original size", RGB(128, 0, 128)) _
)

' === Fill Power-Up Legend ===
For i = LBound(pData) To UBound(pData)
    ws.Cells(pStartRow + i, 30).Value = pData(i)(0) ' AD = Type
    ws.Cells(pStartRow + i, 31).Value = pData(i)(1) ' AE = Description
    ws.Cells(pStartRow + i, 32).Interior.Color = pData(i)(2) ' AF = Colour swatch
    ws.Cells(pStartRow + i, 30).HorizontalAlignment = xlCenter
    ws.Cells(pStartRow + i, 31).HorizontalAlignment = xlLeft
Next i

' === Format Borders Around Legend ===
With ws.Range("AD13:AF" & (pStartRow + UBound(pData)))
    .Borders.LineStyle = xlContinuous
    .Borders.Color = RGB(180, 180, 180)
    .Font.Size = 10
    .RowHeight = boardRowHeight
End With

' === Adjust Column Widths for Neat Layout ===
ws.Columns("AD").ColumnWidth = 10
ws.Columns("AE").ColumnWidth = 30
ws.Columns("AF").ColumnWidth = 8
End Sub

' --- Clear Chaser ---
Sub ClearChaser()
    Dim i As Long
    Dim ub As Long

    'I couldn't figure out the Ubound Error..... :(
    On Error Resume Next
    ub = UBound(chaserR)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        chaserActive = False
        Exit Sub
    End If
    On Error GoTo 0

    For i = 0 To ub
        If chaserR(i) >= 2 And chaserR(i) <= 27 And chaserC(i) >= 2 And chaserC(i) <= 26 Then
            Cells(chaserR(i), chaserC(i)).Interior.Color = boardColor
            Cells(chaserR(i), chaserC(i)).Value = ""
        End If
    Next i

    ' Clear the arrays from memory and disable chaser
    On Error Resume Next
    Erase chaserR
    Erase chaserC
    chaserActive = False

    ' Cancel any scheduled chaser timer
    If nextChaserUpdateTime <> 0 Then
        On Error Resume Next
        Application.OnTime nextChaserUpdateTime, "UpdateChaser", Schedule:=False
        On Error GoTo 0
    End If
End Sub
Private Sub SleepSeconds(seconds As Double)
    Dim endTime As Date
    endTime = Now + seconds / 86400#
    Do While Now < endTime
        DoEvents
    Loop
End Sub
Sub DrawVictoryFace()
    Dim offsets As Variant
    Dim facePath() As Variant
    Dim paintedCells As Object
    Dim i As Long, p As Long, seg As Long
    Dim stepDelay As Double
    Dim nextR As Long, nextC As Long
    Dim targetR As Long, targetC As Long
    Dim rr As Long, cc As Long
    Dim centerRow As Long, centerCol As Long
    Dim startR As Long, startC As Long
    Dim oldTailR As Long, oldTailC As Long
    Dim key As String
    Dim neededLen As Long
    Dim oldBoardColor As Long, oldBodyColor As Long, oldHeadColor As Long

    Const TOP_ROW As Long = 2
    Const BOTTOM_ROW As Long = 27
    Const LEFT_COL As Long = 2
    Const RIGHT_COL As Long = 26
    stepDelay = 0.1

    On Error Resume Next
    oldBoardColor = boardColor
    oldBodyColor = snakeBodyColor
    oldHeadColor = snakeheadColor
    On Error GoTo 0

    boardColor = RGB(0, 0, 0)
    snakeBodyColor = RGB(200, 0, 0)
    snakeheadColor = RGB(255, 80, 80)

    For rr = TOP_ROW To BOTTOM_ROW
        For cc = LEFT_COL To RIGHT_COL
            Cells(rr, cc).Interior.Color = boardColor
            Cells(rr, cc).Value = ""
        Next cc
    Next rr

    offsets = Array( _
        Array(-4, -3), Array(-4, 3), _
        Array(-1, -4), Array(0, -3), Array(1, -2), _
        Array(2, -1), Array(2, 0), Array(2, 1), _
        Array(1, 2), Array(0, 3), Array(-1, 4) _
    )

    centerRow = (TOP_ROW + BOTTOM_ROW) \ 2
    centerCol = (LEFT_COL + RIGHT_COL) \ 2
    ReDim facePath(LBound(offsets) To UBound(offsets))
    For i = LBound(offsets) To UBound(offsets)
        targetR = centerRow + offsets(i)(0)
        targetC = centerCol + offsets(i)(1)
        If targetR < TOP_ROW Then targetR = TOP_ROW
        If targetR > BOTTOM_ROW Then targetR = BOTTOM_ROW
        If targetC < LEFT_COL Then targetC = LEFT_COL
        If targetC > RIGHT_COL Then targetC = RIGHT_COL
        facePath(i) = Array(targetR, targetC)
    Next i

    Set paintedCells = CreateObject("Scripting.Dictionary")

    neededLen = 8
    ReDim r(0 To neededLen)
    ReDim c(0 To neededLen)

    startR = centerRow
    startC = centerCol - 6
    If startC - neededLen < LEFT_COL Then startC = LEFT_COL + neededLen
    If startC > RIGHT_COL Then startC = RIGHT_COL

    For seg = 0 To neededLen
        r(seg) = startR
        c(seg) = startC - seg
    Next seg

    ShowSnake
    DoEvents
    SleepSeconds 0.25

    For p = LBound(facePath) To UBound(facePath)
        targetR = facePath(p)(0)
        targetC = facePath(p)(1)

        Do While r(0) <> targetR Or c(0) <> targetC

            oldTailR = r(UBound(r))
            oldTailC = c(UBound(c))

            nextR = r(0)
            nextC = c(0)
            If c(0) < targetC Then
                nextC = c(0) + 1
            ElseIf c(0) > targetC Then
                nextC = c(0) - 1
            ElseIf r(0) < targetR Then
                nextR = r(0) + 1
            ElseIf r(0) > targetR Then
                nextR = r(0) - 1
            End If

            For seg = UBound(r) To 1 Step -1
                r(seg) = r(seg - 1)
                c(seg) = c(seg - 1)
            Next seg
            r(0) = nextR
            c(0) = nextC

            key = oldTailR & "," & oldTailC
            If Not paintedCells.Exists(key) Then
                If oldTailR >= TOP_ROW And oldTailR <= BOTTOM_ROW _
                   And oldTailC >= LEFT_COL And oldTailC <= RIGHT_COL Then
                    Cells(oldTailR, oldTailC).Interior.Color = boardColor
                End If
            End If

            ShowSnake
            DoEvents
            SleepSeconds stepDelay
        Loop

        If targetR >= TOP_ROW And targetR <= BOTTOM_ROW _
           And targetC >= LEFT_COL And targetC <= RIGHT_COL Then
            Cells(targetR, targetC).Interior.Color = RGB(255, 0, 0)
            paintedCells.Add targetR & "," & targetC, True
        End If

        SleepSeconds 0.08
        DoEvents
    Next p

    ReDim Preserve r(0 To 4)
ReDim Preserve c(0 To 4)
    For seg = 0 To UBound(r)
        If r(seg) >= TOP_ROW And r(seg) <= BOTTOM_ROW _
           And c(seg) >= LEFT_COL And c(seg) <= RIGHT_COL Then
            Cells(r(seg), c(seg)).Interior.Color = RGB(255, 0, 0)
        End If
    Next seg

    With Range("AD21")
        .Value = "** You beat Level 30 — Snake Champion! **"
        .Font.Color = RGB(255, 128, 0)
        .Font.Bold = True
    End With

    On Error Resume Next
    boardColor = oldBoardColor
    snakeBodyColor = oldBodyColor
    snakeheadColor = oldHeadColor
    On Error GoTo 0

End Sub


