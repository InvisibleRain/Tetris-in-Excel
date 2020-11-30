'32-bit and 64-bit systems needs differenct dll functions
#If Win64 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" _
            (ByVal vKey As Long) As Integer
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" _
            (ByVal vKey As Long) As Integer
#End If

            
'Hex values of keys for GetAsyncKeyState
Private Enum Keys
    Down = &H28
    Up = &H26
    Right = &H27
    Left = &H25
    LControl = &HA2
End Enum

'20*10 screen size
Public Const screenWidth = 10 - 1 'that's fucked up and need fix
Public Const screenHeight = 20 - 1 'that's fucked up too
Public screen(screenHeight, screenWidth) As Integer

'Start position of the screen on a worksheet. [1,1] is top left corner
Public Const screenY As Integer = 4
Public Const screenX As Integer = 5

Public Const nextScreenWidth = 4 - 1 'that's fucked up and need fix
Public Const nextScreenHeight = 4 - 1 'that's fucked up too

Public Const nextScreenY As Integer = screenY + 5
Public Const nextScreenX As Integer = screenX + 11

Public nextScreen(nextScreenHeight, nextScreenWidth) As Integer 'screen of the next tetromino preview

'Allow player to control tetromino
Public Sub GiveControlls(tetr As Object)
'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    If GetAsyncKeyState(Down) Then
        Call tetr.Move("down")
    ElseIf GetAsyncKeyState(Up) Then
        Call tetr.Rotate("right")
    ElseIf GetAsyncKeyState(Left) Then
        Call tetr.Move("left")
    ElseIf GetAsyncKeyState(Right) Then
        Call tetr.Move("right")
    ElseIf GetAsyncKeyState(LControl) Then
        Call tetr.Move("drop")
    End If
    
    Application.ScreenUpdating = True
End Sub

'Removes tetr from screen() array
Public Sub RemoveFromScreen(ByRef screen() As Integer, ByRef tetr As Object)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False

    Dim col As Integer
    Dim row As Integer
    
      For row = 0 To tetr.height - 1
        For col = 0 To tetr.width - 1
            'Tetromono structure and color stored separately in object
            If tetr.FigureCell(row, col) <> 0 Then
                screen(row + tetr.PosY, col + tetr.PosX) = 0
            End If
            
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

Public Sub PutOnScreen(ByRef screen() As Integer, ByRef tetr As Object)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False

    Dim col As Integer
    Dim row As Integer
    
      For row = 0 To tetr.height - 1
    
        For col = 0 To tetr.width - 1
            'Tetromono structure and color stored separately in object
            If tetr.FigureCell(row, col) <> 0 Then 'if branch just for optimization
                screen(row + tetr.PosY, col + tetr.PosX) = tetr.color
            End If
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

'Kids, don't do this at home
Public Sub PutOnNextScreen(ByRef screen() As Integer, ByRef tetr As Object)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False

    Dim col As Integer
    Dim row As Integer
    
      For row = 0 To tetr.height - 1
    
        For col = 0 To tetr.width - 1
            'Tetromono structure and color stored separately in object
            If tetr.FigureCell(row, col) <> 0 Then 'if branch just for optimization
                screen(row, col) = tetr.color
            End If
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

Public Sub DrawScreen(screen() As Integer)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
   'screenY and screenX is start of screen that defined in constants on top of this module
    For row = 0 To screenHeight
        For col = 0 To screenWidth
            If (screen(row, col) <> 0) Then
                With Worksheets("Sheet1").Cells(row + screenY, col + screenX)
                 .Interior.ColorIndex = screen(row, col)
                End With
            Else
                With Worksheets("Sheet1").Cells(row + screenY, col + screenX)
                 If (.Interior.ColorIndex <> xlNone) Then
                    .Interior.ColorIndex = xlNone
                    .Borders(xlEdgeRight).LineStyle = xlNone
                    .Borders(xlEdgeLeft).LineStyle = xlNone
                    .Borders(xlEdgeBottom).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlNone
                 End If
                End With
            End If
        Next col
    Next row
    
    
    
    Call Engine.AddOutline(Engine.screen) 'Duct tape
    
    Application.ScreenUpdating = True
End Sub

'Quick and dirty hack. It would be better, if screen was an object with predefined coordinates and size
Public Sub DrawNextScreen(screen() As Integer)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
   'screenY and screenX is start of screen that defined in constants on top of this module
    For row = 0 To nextScreenHeight
        For col = 0 To nextScreenWidth
            If (screen(row, col) <> 0) Then
                With Worksheets("Sheet1").Cells(row + nextScreenY, col + nextScreenX)
                 .Interior.ColorIndex = screen(row, col)
                End With
            Else
                With Worksheets("Sheet1").Cells(row + screenY, col + screenX)
                 If (.Interior.ColorIndex <> xlNone) Then
                    .Interior.ColorIndex = xlNone
                    .Borders(xlEdgeRight).LineStyle = xlNone
                    .Borders(xlEdgeLeft).LineStyle = xlNone
                    .Borders(xlEdgeBottom).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlNone
                 End If
                End With
            End If
        Next col
    Next row
    
    Call Engine.AddOutlineToNextScreen(Engine.nextScreen) 'Duct tape
    
    Application.ScreenUpdating = True
End Sub

'Adds outline to every block on screen
'Separate procedure is needed to fix bug
Public Sub AddOutline(screen() As Integer)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
    For row = 0 To screenHeight
        For col = 0 To screenWidth
            If (screen(row, col) <> 0) Then
                With Worksheets("Sheet1").Cells(row + screenY, col + screenX)
                 .Borders(xlEdgeRight).LineStyle = xlContinuous
                 .Borders(xlEdgeLeft).LineStyle = xlContinuous
                 .Borders(xlEdgeBottom).LineStyle = xlContinuous
                 .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
            End If
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

Public Sub AddOutlineToNextScreen(screen() As Integer)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
    For row = 0 To nextScreenHeight
        For col = 0 To nextScreenWidth
            If (screen(row, col) <> 0) Then
                With Worksheets("Sheet1").Cells(row + nextScreenY, col + nextScreenX)
                 .Borders(xlEdgeRight).LineStyle = xlContinuous
                 .Borders(xlEdgeLeft).LineStyle = xlContinuous
                 .Borders(xlEdgeBottom).LineStyle = xlContinuous
                 .Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
            End If
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

Public Sub ClearScreen(Optional tetr As Object) 'idk why tetr is here
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Dim col As Integer
    Dim row As Integer
    
    For row = screenY To screenHeight + screenY
        For col = screenX To screenWidth + screenX
            If (Worksheets("Sheet1").Cells(row, col).Interior.color <> xlNone) Then 'if branch just for optimization, but it will lag when too many cells on the screen
                With Worksheets("Sheet1").Cells(row, col)
                 .Interior.ColorIndex = xlNone
                 .Borders(xlEdgeRight).LineStyle = xlNone
                 .Borders(xlEdgeLeft).LineStyle = xlNone
                 .Borders(xlEdgeBottom).LineStyle = xlNone
                 .Borders(xlEdgeTop).LineStyle = xlNone
                End With
            End If
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

'another quick and dirty hack
Public Sub ClearNextScreen()
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Const White As Integer = 2
    
    Dim col As Integer
    Dim row As Integer
    
    For row = nextScreenY To nextScreenHeight + nextScreenY
        For col = nextScreenX To nextScreenWidth + nextScreenX
            If (Worksheets("Sheet1").Cells(row, col).Interior.color <> xlNone) Then 'if branch just for optimization, but it will lag when too many cells on the screen
                With Worksheets("Sheet1").Cells(row, col)
                 .Interior.ColorIndex = White
                 .Borders(xlEdgeRight).LineStyle = xlNone
                 .Borders(xlEdgeLeft).LineStyle = xlNone
                 .Borders(xlEdgeBottom).LineStyle = xlNone
                 .Borders(xlEdgeTop).LineStyle = xlNone
                End With
            End If
        Next col
    Next row
    
    For row = 0 To nextScreenHeight
        For col = 0 To nextScreenWidth
            Engine.nextScreen(row, col) = 0
        Next col
    Next row
    
    Application.ScreenUpdating = True
End Sub

Public Sub ClearScreenArray()
    Dim col As Integer
    Dim row As Integer
    
    For row = 0 To screenHeight
        For col = 0 To screenWidth
            Engine.screen(row, col) = 0
        Next col
    Next row
End Sub

Public Function rnd_num(lowerbound As Integer, upperbound As Integer) As Integer
   Randomize
   rnd_num = Int((upperbound - lowerbound + 1) * rnd + lowerbound)
End Function

