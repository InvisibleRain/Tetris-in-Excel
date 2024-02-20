Private Enum Colors
'Colors value is excel colorindex
 red = 3
 green = 4
 blue = 5
 yellow = 27
 cyan = 8
 purple = 18
 orange = 45
End Enum

Private p_Color As Integer
Private p_PosX As Integer
Private p_PosY As Integer
Private p_width As Integer
Private p_height As Integer
Private p_Figure() As Integer


Public Property Get color() As Integer
    color = p_Color
End Property

'Private Property Let color(value As Integer)
'    p_Color = value
'End Property

Public Property Get PosX() As Integer
    PosX = p_PosX
End Property

Public Property Let PosX(value As Integer)
    p_PosX = value
End Property

Public Property Get PosY() As Integer
    PosY = p_PosY
End Property

Public Property Let PosY(value As Integer)
    p_PosY = value
End Property

Public Property Get width() As Integer
    width = p_width
End Property

Public Property Get height() As Integer
    height = p_height
End Property

Public Property Get Figure() As Integer()
    Figure = p_Figure
End Property

Public Property Get FigureCell(row, col As Integer) As Integer
    FigureCell = p_Figure(row, col)
End Property

Private Property Let FigureCell(row, col As Integer, value As Integer)
    p_Figure(row, col) = value
End Property

Private Sub Class_Initialize()
'This tons of lines is holy duct tape for replace initialization of variables with declaration (that isn't exist in VBA)
'iFigure is "stick", oFigure is square brick, and so on
Static iFigure(0, 3) As Integer
Static ofigure(1, 1) As Integer
Static sfigure(1, 2) As Integer
Static zfigure(1, 2) As Integer
Static jfigure(1, 2) As Integer
Static Lfigure(1, 2) As Integer
Static tfigure(1, 2) As Integer
   '---------------'
   iFigure(0, 0) = 1
   iFigure(0, 1) = 1
   iFigure(0, 2) = 1
   iFigure(0, 3) = 1

   '---------------'
   ofigure(0, 0) = 1
   ofigure(0, 1) = 1
   
   ofigure(1, 0) = 1
   ofigure(1, 1) = 1
    
   '---------------'
   sfigure(0, 0) = 0
   sfigure(0, 1) = 1
   sfigure(0, 2) = 1
   
   sfigure(1, 0) = 1
   sfigure(1, 1) = 1
   sfigure(1, 2) = 0
    
   '---------------'
   zfigure(0, 0) = 1
   zfigure(0, 1) = 1
   zfigure(0, 2) = 0
   
   zfigure(1, 0) = 0
   zfigure(1, 1) = 1
   zfigure(1, 2) = 1
    
   '---------------'
   jfigure(0, 0) = 1
   jfigure(0, 1) = 0
   jfigure(0, 2) = 0
   
   jfigure(1, 0) = 1
   jfigure(1, 1) = 1
   jfigure(1, 2) = 1
    
   '---------------'
   Lfigure(0, 0) = 0
   Lfigure(0, 1) = 0
   Lfigure(0, 2) = 1
   
   Lfigure(1, 0) = 1
   Lfigure(1, 1) = 1
   Lfigure(1, 2) = 1
    
   '---------------'
   tfigure(0, 0) = 0
   tfigure(0, 1) = 1
   tfigure(0, 2) = 0
   
   tfigure(1, 0) = 1
   tfigure(1, 1) = 1
   tfigure(1, 2) = 1
   
'End of duct tape

'Creating random figure
Randomize
Select Case Engine.rnd_num(1, 7)
   Case 1
       ReDim Preserve p_Figure(UBound(iFigure, 1), UBound(iFigure, 2))
       p_Figure = iFigure
       p_Color = cyan
   Case 2
       ReDim Preserve p_Figure(UBound(ofigure, 1), UBound(ofigure, 2))
       p_Figure = ofigure
       p_Color = yellow
   Case 3
       ReDim Preserve p_Figure(UBound(sfigure, 1), UBound(sfigure, 2))
       p_Figure = sfigure
       p_Color = green
   Case 4
       ReDim Preserve p_Figure(UBound(zfigure, 1), UBound(zfigure, 2))
       p_Figure = zfigure
       p_Color = red
   Case 5
       ReDim Preserve p_Figure(UBound(jfigure, 1), UBound(jfigure, 2))
       p_Figure = jfigure
       p_Color = blue
   Case 6
       ReDim Preserve p_Figure(UBound(Lfigure, 1), UBound(Lfigure, 2))
       p_Figure = Lfigure
       p_Color = orange
   Case 7
       ReDim Preserve p_Figure(UBound(tfigure, 1), UBound(tfigure, 2))
       p_Figure = tfigure
       p_Color = purple
    End Select
   
    p_width = CalculateWidth()
    p_height = CalculateHeight()
       
    p_PosY = 0
    p_PosX = 3
End Sub

Private Function CalculateWidth() As Integer
    CalculateWidth = UBound(p_Figure, 2) - LBound(p_Figure, 2) + 1
End Function

Private Function CalculateHeight() As Integer
    CalculateHeight = UBound(p_Figure, 1) - LBound(p_Figure, 1) + 1 'num of rows
End Function

Private Function CanRotate(direction As String) As Boolean
    Select Case direction
        Case "left"
            If (Me.PosX - (height - 1)) >= 0 Then
            CanRotate = True
            End If
        Case "right"
            If (Me.PosX + (height - 1)) <= Engine.screenWidth Then
            CanRotate = True
            End If
        Case Else
            Err.Raise 666, , "Invalid Argument In tetr.CanRotate Method." & vbCrLf & "You can only use ""left"" and ""right"" strings (with lower letter)"
    End Select
End Function

Public Sub Rotate(direction As String)
    If (direction = "left" And CanRotate("left")) Or (direction = "right" And CanRotate("right")) Then
        Call Engine.RemoveFromScreen(Engine.screen(), Me)
        
        'figureCopy is needed to not affect itself while rotating
        Dim figureCopy() As Integer
        ReDim Preserve figureCopy(Me.height - 1, Me.width - 1) 'We can't declare the array in the line above, because VBA requires a "constant"
        figureCopy() = Me.Figure()
        ReDim p_Figure(Me.width - 1, Me.height - 1)
        
        maxCol = UBound(figureCopy, 2)
        maxRow = UBound(figureCopy, 1)
        
        Dim row As Integer
        Dim col As Integer
        
        For row = 0 To maxRow
            For col = 0 To maxCol
                Select Case direction
                    Case "left"
                    FigureCell(maxCol - col, row) = figureCopy(row, col)
                    Case "right"
                    FigureCell(col, maxRow - row) = figureCopy(row, col)
                    Case Else
                    Err.Raise 666, , "Invalid Argument In tetr.Rotate Method." & vbCrLf & "You can only use ""left"" and ""right"" strings (with lower letter)"
                End Select
            Next col
        Next row
        
        p_width = CalculateWidth()
        p_height = CalculateHeight()
        
        Call Engine.PutOnScreen(Engine.screen(), Me)
        Call Engine.DrawScreen(Engine.screen)
    End If
End Sub

Public Function CanMove(direction As String) As Boolean
    Dim row As Integer
    Dim col As Integer
    
    Select Case direction
        Case "left"
        If (Me.PosX > 0) Then
            For row = 0 To Me.height - 1
                For col = 0 To Me.width - 1
                    If (Me.FigureCell(row, col) <> 0) Then
                        If (Engine.screen(Me.PosY + row, Me.PosX + col - 1) <> 0) Then
                            CanMove = False
                            Exit Function
                        End If
                        Exit For
                    End If
                Next col
            Next row
            CanMove = True
        Else
            CanMove = False
        End If
            
        Case "right"
        If (Me.PosX + Me.width - 1 < Engine.screenWidth) Then
            For row = 0 To Me.height - 1
                For col = Me.width - 1 To 0 Step -1
                    If (Me.FigureCell(row, col) <> 0) Then
                        If (Engine.screen(Me.PosY + row, Me.PosX + col + 1) <> 0) Then
                            CanMove = False
                            Exit Function
                        End If
                        Exit For
                    End If
                Next col
            Next row
            CanMove = True
        Else
            CanMove = False
        End If
            
        Case "down"
        'Finding lowest cell in each collumn and compare it to screen cell that under lowest cell
        'If both lowest cell and screnn cell not empty then return false
        If (Me.PosY + Me.height - 1 < Engine.screenHeight) Then
            For col = 0 To Me.width - 1
                For row = Me.height - 1 To 0 Step -1
                    If (Me.FigureCell(row, col) <> 0) Then
                        If (Engine.screen(Me.PosY + row + 1, Me.PosX + col) <> 0) Then
                            CanMove = False
                            Exit Function
                        End If
                        Exit For
                    End If
                Next row
            Next col
            CanMove = True
        Else
            CanMove = False
        End If
        
        Case Else
        Err.Raise 666, , "Invalid Argument In tetr.CanMove Method." & vbCrLf & "You can only use ""left"", ""right"" and ""down"" strings (with lower letter)"
        
        End Select
End Function

Public Sub Move(direction As String)
    'This line for optimization. Updating will be turned on in the end of procedure
    Application.ScreenUpdating = False
    
    Call Engine.RemoveFromScreen(Engine.screen(), Me)
    
    'Temporary duct tape, I will add isThereSomething(direction) method
    Select Case direction
        Case "left"
            If (Me.CanMove("left")) Then
            Let Me.PosX = Me.PosX - 1
            End If
        Case "right"
            If (Me.CanMove("right")) Then
            Let Me.PosX = Me.PosX + 1
            End If
        Case "down"
            If (Me.CanMove("down")) Then
            Let Me.PosY = Me.PosY + 1
            End If
        Case "drop"
            While (Me.CanMove("down"))
            Let Me.PosY = Me.PosY + 1
            Wend
        Case Else
            Err.Raise 666, , "Invalid Argument In tetr.Move Method." & vbCrLf & "You can only use ""left"", ""right"" and ""down"" strings (with lower letter)"
        End Select
        
    
    Call Engine.PutOnScreen(Engine.screen(), Me)
    Call Engine.DrawScreen(Engine.screen)
    
    Application.ScreenUpdating = True
End Sub

