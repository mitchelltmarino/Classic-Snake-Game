Attribute VB_Name = "UtilityModule"
'==== Snake Game by Mitchell Marino ========================+'
' Name: Mitchell Marino
' Date: 03/04/2017
' Program title: SnakeGame.xlsm
' Description: An assortment of utilities for functionality.
'===========================================================+'
Option Explicit

'Initialization.
Public Sub Initialize_Game()

    DeclareKeys
    Session_FilePath = ""
    Game_Setup

End Sub

'Declare keys.
Public Sub DeclareKeys()
    Application.OnKey "{UP}", "PressedUp"
    Application.OnKey "{DOWN}", "PressedDown"
    Application.OnKey "{LEFT}", "PressedLeft"
    Application.OnKey "{RIGHT}", "PressedRight"
    Application.OnKey " ", "Game_Start"
    Application.ScreenUpdating = True
End Sub

'Open high scores.
Sub HighScoresFrame()
    'If file path not yet set,
    If Session_FilePath = "" Then
        'Request user to set the file path.
        MsgBox ("Please select the high scores database file for the current session")
        'Get the file path via function.
        Session_FilePath = Get_File_Path
    End If
    'If filepath was set,
    If Session_FilePath <> "" Then
        'Show highscores.
        frmHighScores.Show
    End If
End Sub

'Open user frame for user to enter info so they can input their score to high scores.
Sub EnterScoresFrame()

    'If filepath not set yet,
    If Session_FilePath = "" Then
        'Request user to set file path.
        MsgBox ("Please select the high scores database file for the current session")
        'Get file path via function.
        Session_FilePath = Get_File_Path
    End If
    'If file path was set,
    If Session_FilePath <> "" Then
        'Open userframe.
        frmScoreInput.Show
    End If
End Sub

'Used to set file path.
Private Function Get_File_Path() As String
    
    'Opening a file dialog.
    Dim fd As Office.FileDialog
    Dim filepath As String
    
    'Define file dialog.
    Set fd = Application.FileDialog(msoFileDialogOpen)
    'Initial file name.
    fd.InitialFileName = ThisWorkbook.Path
    
    'Open file dialog.
    With fd
        If .Show Then
            'Set the file path String.
            filepath = .SelectedItems(1)
        End If
    End With
    
    If Right(UCase(filepath), 4) <> ".MDB" Then
        MsgBox "Database file name requires valid extension, '.xlsx'", vbCritical, "Error"
        filepath = ""
    End If
    
    'Return file path as Get_File_Path
    Get_File_Path = filepath
    
End Function
