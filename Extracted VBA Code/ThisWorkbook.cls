VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==== Snake Game by Mitchell Marino ========================+'
' Name: Mitchell Marino
' Date: 03/04/2017
' Program title: SnakeGame.xlsm
' Description: Runs code on startup... To initialize keys
'===========================================================+'
Option Explicit

Public Sub Workbook_Open()
    'On workbook open, declare keys using utility module function.
    DeclareKeys
    'Filepath initially empty.
    Session_FilePath = ""
    'Set up the game.
    Game_Setup

End Sub

