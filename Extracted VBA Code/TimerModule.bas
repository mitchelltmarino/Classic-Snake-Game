Attribute VB_Name = "TimerModule"
'==== Snake Game by Mitchell Marino ========================+'
' Name: Mitchell Marino
' Date: 03/04/2017
' Program title: SnakeGame.xlsm
' Description: Module used to manage the timer.
'===========================================================+'
Option Explicit

'Set timer function from user32 library.
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, _
ByVal nIDEvent As LongPtr, _
ByVal uElapse As LongPtr, _
ByVal lpTimerFunc As LongPtr) As Long
'Killtimer function from user32 library.
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, _
ByVal nIDEvent As LongPtr) As LongPtr
'TimerID as a long.
Private TimerID As Long
'Sleep function from kernel32 library.
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

 
 'Duration for SetTimer is in miliseconds.
Public Sub StartTimer()
    'Select the top left cell.
    SnakeGame.Range("A1").Select
    'Enable screenupdating.
    Application.ScreenUpdating = True
    'Initially sleep for half a second.
    Sleep (500)
    
    'If the timer isn't already running, start it.
    If TimerID = 0 Then
            'Start a timer that will tick and run TimerEvent every 200 ms.
            TimerID = SetTimer(0, 0, 200, AddressOf TimerEvent)
            'If timer not initialized, then it failed.
            If TimerID = 0 Then
                MsgBox "Timer not created."
        End If
    Else
        'Stop timer if it already is on.
        StopTimer
    End If
End Sub

'StopTimer.
Public Sub StopTimer()
     'If the timer is already running, shut it off.
    If TimerID <> 0 Then
        'Kill the timer.
        KillTimer 0, TimerID
        'TimerID is once again 0.
        TimerID = 0
    End If
End Sub
 
'Timer event, is triggered every time timer ticks.
Private Sub TimerEvent()
    
    'Handle errors.
    On Error Resume Next
    'Run the game. (1 tick worth)
    Game_Run
    
End Sub



