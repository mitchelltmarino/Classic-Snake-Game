Attribute VB_Name = "MainModule"
'==== Snake Game by Mitchell Marino ========================+'
' Name: Mitchell Marino
' Date: 03/04/2017
' Program title: SnakeGame.xlsm
' Description: Main module for the Snake Game.
'===========================================================+'
Option Explicit

'Public Variables.
Public snake() As Range  'Array holding snake range.
                         'Snake(0) = head.
Public Score As Integer  'Score the user has obtained.
Public Session_FilePath As String 'Filepath for the scores of the current session.
                       
'Private Variables.
Private keyRefresh As Boolean     'Global cooldown on key press. (So new direction does not change before intended action has occurred)
Private snake_Direction As String 'Direction snake is facing.
Private snake_ToGrow As Boolean   'Boolean to track if snake will grow out the back on next movement.
Private snake_Length As Integer   'Snake Length.
Private Map_Range As Range        'Map bounds.

'Initialize values for the game. (Set it up)
Public Sub Game_Setup()

    'Set Snake Direction.
    snake_Direction = "UP"

    'Key is ready to be pressed.
    keyRefresh = True
    'Snake will not grow on first movement.
    snake_ToGrow = False

    'Set map range.
    Set Map_Range = Range("B2:AE31")
    'Set map interior color to white space.
    Map_Range.Interior.Color = RGB(255, 255, 255)

    'Initialize snake_Length to 3.
    snake_Length = 3

    'Set the dimensions fo snake to 3.
    ReDim snake(3)
    
    'Original snake starting point, and values. (length of 3)
    Set snake(0) = Range("P15")
    Set snake(1) = Range("P16")
    Set snake(2) = Range("P17")
    
    'Head of snake blue, body of snake black.
    Range("P15").Interior.Color = RGB(0, 0, 255)
    Range("P16").Interior.Color = RGB(0, 0, 0)
    Range("P17").Interior.Color = RGB(0, 0, 0)

End Sub

'Start Game.
Public Sub Game_Start()
    'Set up the game.
    Call Game_Setup
    'Generate new apple on the map.
    New_Food
    'Stop timer if it was going for some reason.
    TimerModule.StopTimer
    'Start timer.
    TimerModule.StartTimer 'Program will now run according to the ticking of the timer!
    
End Sub

'Generate new apple on random location.
Public Sub New_Food()
    
    'Cell is temporary location.
    Dim cell As Range
    'Generate random location within map bounds.
    Set cell = Map_Range.Cells(Int(Rnd * Map_Range.Cells.Count) + 1)
    
    'If cell is not on empty space,
    Do Until cell.Interior.Color = RGB(255, 255, 255)
        'Generate another location until it is.
        Set cell = Map_Range.Cells(Int(Rnd * Map_Range.Cells.Count) + 1)
    Loop
    
    'Set the location of the apple to be red.
    cell.Interior.Color = RGB(255, 0, 0)

End Sub

'Game's main loop.
Public Sub Game_Run()
    'Move snake.
    Call Snake_Move

End Sub


'Moves the snake and updates its values.
Public Sub Snake_Move()
    
    'Tracks the index of the snake body.
    Dim i As Integer
    
    'Temp location used to store end of snake if need be.
    Dim tempLocation As Range
        
    'If snake ate an apple and will grow on this movement..
    If snake_ToGrow = True Then
        'Store last space.
        Set tempLocation = snake(snake_Length - 1)
    'Otherwise,
    Else
        'Last space will become whitespace.
        snake(snake_Length - 1).Interior.Color = RGB(255, 255, 255)
    End If
        
    'Start at end of snake.
   For i = snake_Length - 1 To 1 Step -1
        'Shift, starting at the end of the snake, the latter block into the prior for each pair.
       Set snake(i) = snake(i - 1)
   Next
   
   'If snake is set to grow,
   If snake_ToGrow = True Then
        'Now increase length of snake by 1.
        snake_Length = snake_Length + 1
        'Redimension the array to the new size.
        ReDim Preserve snake(snake_Length)
        'Set the snakes last location to the previous last location.
        Set snake(snake_Length - 1) = tempLocation
   End If
    
    'Snake not to grow again.
    snake_ToGrow = False
    
    'Snake moves based on current way it is facing.
    If snake_Direction = "UP" Then
        Set snake(0) = snake(0).Offset(-1, 0) 'Move up 1.
    
    ElseIf snake_Direction = "DOWN" Then
        Set snake(0) = snake(0).Offset(1, 0) 'Move Down 1.
        
    ElseIf snake_Direction = "LEFT" Then
        Set snake(0) = snake(0).Offset(0, -1) 'Move Left 1.
        
    Else 'If snake_Direction = "RIGHT" Then
        Set snake(0) = snake(0).Offset(0, 1) 'Move Right 1.
        
    End If

    'Check for snake interaction.
    If snake(0).Interior.Color = RGB(255, 0, 0) Then 'If snake has stumbled into an apple,
        snake_ToGrow = True 'Snake will grow on next movement.
        New_Food            'Generate new apple.
    ElseIf snake(0).Interior.Color <> RGB(255, 255, 255) Then 'If the snake has ran into itself or a wall,
        'Collision occurred.
        Range("AG7").Value = snake_Length 'Store the snake's length.
        TimerModule.StopTimer             'Stop the timer.
        'Notify the user on what has occurred. (Collision, score and how they can save their score)
        MsgBox "You lost because you collided with something." & vbCrLf _
        & "Your final length was " & snake_Length & "." & vbCrLf _
        & "To save your score, press the submit high score button next to the score meter."
        'Exit the subroutine.
        Exit Sub
    End If
    
    'New key is ready to be pressed.
    keyRefresh = True
    
        'Interior colour for head = blue.
        snake(0).Interior.Color = RGB(0, 0, 255)
        'Interior colour for neck (2nd to head) is black now.
        snake(1).Interior.Color = RGB(0, 0, 0)
    
    'Update the score field with the snake's current length.
    Range("AG7").Value = snake_Length

End Sub
    
'If up button is pressed,
Public Sub PressedUp()

    'If snake direction is not down, and key is available.
    If snake_Direction <> "DOWN" And keyRefresh = True Then
        'Direction is now up.
        snake_Direction = "UP"
    End If

    'Key no longer available.
    keyRefresh = False

End Sub

'If down button is pressed,
Public Sub PressedDown()
    
    'If snake direction is not up, and key is available.
    If snake_Direction <> "UP" And keyRefresh = True Then
        snake_Direction = "DOWN"
    End If

    'Key no longer available.
    keyRefresh = False

End Sub

'If left button is pressed,
Public Sub PressedLeft()

    'If snake direction is not right, and key is available,
    If snake_Direction <> "RIGHT" And keyRefresh = True Then
        'Direction is now right.
        snake_Direction = "LEFT"
    End If

    'Key no longer available.
    keyRefresh = False

End Sub

'If right button is pressed,
Public Sub PressedRight()

    'If snake direction is not left, and key is available,
    If snake_Direction <> "LEFT" And keyRefresh = True Then
        'Direction is now left.
        snake_Direction = "RIGHT"
    End If
    
    'Key no longer available.
    keyRefresh = False

End Sub



