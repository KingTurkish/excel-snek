# Snek in Excel VBA
This version of the game is written in Microsoft® Excel® for Microsoft® Excel® for Microsoft 365 MSO (Version 2508 Build 16.0.19127.20314) 64-bit 

![Screenshot](/images/Excel-Snek.jpg)

Excel Snek Game

Original concept and code by Haronkar Singh (Excel-Snake)
https://github.com/haronkar/excel-snake

Adapted and improved by Tristan Caldwell (Excel-Snek)
https://github.com/KingTurkish/Excel-Snek

Improvements for v1.0 of Excel-Snek

Dynamic high score table

Power ups (Speed, slow, wall pass, traps, shrink, super shrink, bonus)

Colour legends for power ups

Cheat codes

Moving wall levels

Hidden trap levels

Snek Chaser (yes it needs some work)


To customize the looks of the game you can edit the RGB colours in the code

StartGame() sub under ' --- Colours ---

```
 ' --- Colours ---
    boardColor = RGB(255, 246, 211)
    snakeBodyColor = RGB(249, 168, 117)
    snakeheadColor = RGB(235, 107, 111)
    appleColor = RGB(0, 112, 192)
    wallColor = RGB(64, 64, 64)
```

To customize the levels of the game you can edit Level Progression settings in the code

```
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
```
