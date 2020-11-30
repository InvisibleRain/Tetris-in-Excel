Private GameState As Byte
Enum GameStates
    Stopped = 0
    Running = 1
    Paused = 2
End Enum

Public score As Integer

Private tetr As Tetromino
Private nextTetr As Tetromino

Sub Game()
    'Each cycle - one tetromino
    Do While (GameState <> Stopped)
    
        If nextTetr Is Nothing Then
            Set tetr = New Tetromino
            Set nextTetr = New Tetromino
        Else
            Set tetr = nextTetr
            Set nextTetr = New Tetromino
        End If
        
        Call Engine.PutOnNextScreen(Engine.nextScreen, nextTetr)
        Call Engine.DrawNextScreen(Engine.nextScreen)
        
        Call DisplayScore(score)
        
        Do 'Each cycle - one frame
            Call PauseIfPaused
            
            'Look in ResetButton_Click sub for explanation
            If GameState = Stopped Then
                score = 0
                Call DisplayScore(score)
                Call Engine.ClearScreen
                Call Engine.ClearScreenArray
                Set tetr = Nothing
                Exit Sub
            End If
            
            'This line for optimization. Updating will be turned on in the end of procedure
            Application.ScreenUpdating = False
            
            Call Engine.RemoveFromScreen(Engine.screen(), tetr) 'Removes collisions with further frames. Don't move this at the end of the cycle because this tends to bugs
            Call Engine.PutOnScreen(Engine.screen(), tetr)
            Call Engine.DrawScreen(Engine.screen)
            
            Call Engine.GiveControlls(tetr)
            Call tetr.Move("down")
            
            DoEvents
        
            Application.ScreenUpdating = True
        Loop While (tetr.CanMove("down"))
        
        Call CheckGameOver
        Call Engine.PutOnScreen(Engine.screen, tetr)
        Call DestructFilledRows
        Call Engine.DrawScreen(Engine.screen)
        Call ClearNextScreen
        
        If GameState = Stopped Then
            Call Engine.ClearScreen
            Call Engine.ClearScreenArray
            Call ClearNextScreen
            Set tetr = Nothing
            Set nextTetr = Nothing
            Worksheets("Sheet1").Buttons("StartButton").Text = "Start"
            Exit Sub
        End If
        
        Set tetr = Nothing
    Loop
End Sub

Private Sub PauseIfPaused()
    While (GameState = Paused)
        DoEvents
    Wend
End Sub

Public Sub ResetButton_Click()

    If GameState = Stopped Then
        'In case game alredy running Engine.ClearScreen will called from Game Sub
        Call Engine.ClearScreen
        Call Engine.ClearScreenArray
        Call Engine.ClearNextScreen
        Set tetr = Nothing
        Set nextTetr = Nothing
    End If
    
    Call Engine.ClearNextScreen
    Set nextTetr = Nothing
    
    score = 0
    Call DisplayScore(score)
    
    GameState = Stopped
    Worksheets("Sheet1").Buttons("StartButton").Text = "Start"
    
End Sub

Sub StartButton_Click()
    'Changing name of button
    Select Case (Worksheets("Sheet1").Buttons("StartButton").Text)
        Case "Start"
            Worksheets("Sheet1").Buttons("StartButton").Text = "Pause"
            GameState = Running
            score = 0
            Call DisplayScore(score)
            Call Game
            
        Case "Pause"
            Worksheets("Sheet1").Buttons("StartButton").Text = "Resume"
            GameState = Paused
            
        Case "Resume"
            Worksheets("Sheet1").Buttons("StartButton").Text = "Pause"
            GameState = Running
    End Select
End Sub


Private Sub DestructFilledRows()
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
    Dim NumOfFilled As Integer 'Number of non-zero cells in row
    Dim NumOfDestructed As Integer 'Counter. Decision about how much score will be based on this number
    
    For row = 0 To Engine.screenHeight
        NumOfFilled = 0
        
        For col = 0 To Engine.screenWidth
            If Engine.screen(row, col) = 0 Then
                Exit For 'Skip row, if it has at least one empty cell
            Else
                NumOfFilled = NumOfFilled + 1
            End If
        Next col
        
        If NumOfFilled = Engine.screenWidth + 1 Then
            Call DestructRow(row)
            NumOfDestructed = NumOfDestructed + 1
        End If
    Next row
    
    Select Case NumOfDestructed 'Original BPS scoring system
        Case 1
            score = score + 40
        Case 2
            score = score + 100
        Case 3
            score = score + 300
        Case 4
            score = score + 1200
    End Select
    Call DisplayScore(score)
    
    Application.ScreenUpdating = True
End Sub

Private Sub DestructRow(rowNum As Integer)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
    'Destruct row
        For col = 0 To Engine.screenWidth
            Engine.screen(rowNum, col) = 0
        Next col
        
    'Moves blocks above this row down by one step
        For row = rowNum To 1 Step -1
            For col = 0 To Engine.screenWidth
                Engine.screen(row, col) = Engine.screen(row - 1, col)
            Next col
        Next row
    Application.ScreenUpdating = True
End Sub

Private Sub DisplayScore(ByRef score As Integer)
    If score <= 9999999 Then
        Worksheets("Sheet1").Cells((Engine.screenY - 1) + 3, (Engine.screenX - 1) + 12).value = score
    Else
        Worksheets("Sheet1").Cells((Engine.screenY - 1) + 3, (Engine.screenX - 1) + 12).value = ":D"
    End If
End Sub

Private Sub CheckGameOver()
    Dim row As Integer
    Dim col As Integer
    
    row = 0
    For col = 0 To Engine.screenWidth
        If Engine.screen(row, col) <> 0 Then
            GameState = Stopped
            MsgBox "Game Over"
            Exit For
        End If
    Next col
End Sub
