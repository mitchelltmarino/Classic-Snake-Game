VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScoreInput 
   Caption         =   "High Score Input"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9330.001
   OleObjectBlob   =   "frmScoreInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmScoreInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==== Snake Game by Mitchell Marino ========================+'
' Name: Mitchell Marino
' Date: 03/04/2017
' Program title: SnakeGame.xlsm
' Description: Score input frame functionality.
'===========================================================+'
Option Explicit

Private Sub cmdExit_Click()
    'If Exit is clicked, exit the userform.
    Unload Me
End Sub

'Adds a record into the database table.
Private Sub cmdUpdate_Click()
    'Connection variable.
    Dim conn As ADODB.Connection
    'Query String variable.
    Dim strSql As String
    
    'ErrHandler will handler errors.
    On Error GoTo ErrHandler
    
    'If input is valid,
    If ValidInput Then
        'Open new connection.
        Set conn = New ADODB.Connection
        With conn
            .ConnectionString = "Data Source=" & Session_FilePath
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Open
        End With
        
        'Insert the new record.
        conn.Execute "INSERT INTO [Scores] (FirstName, LastName, Score) VALUES ('" & txtFirstName.Text & "', '" & txtLastName & "', " & CInt(SnakeGame.Range("AG7").Value) & ")"
        'Close connection.
        conn.Close
        
        'Connection is now nothing.
        Set conn = Nothing
        txtFirstName.Text = vbNullString
        txtLastName.Text = vbNullString
        'Confirm to user that high scores has been updated.
        MsgBox "High scores has been updated!", vbInformation
    End If
    
'If error occurs, handle it.
ErrHandler:
    If err.Number <> 0 Then
        MsgBox err.Description, vbExclamation, "Error " & CStr(err.Number)
    End If
End Sub

'Determine if input is valid.
Private Function ValidInput()
    ValidInput = False
    
    'As long as first and last name are not numeric, input is considered valid.
    If txtFirstName = "" Or IsNumeric(txtFirstName) = True Then
        MsgBox "Please enter a valid name."
    ElseIf txtLastName = "" Or IsNumeric(txtLastName) = True Then
        MsgBox "Please enter a valid name."
    Else
        ValidInput = True
    End If
    
End Function

