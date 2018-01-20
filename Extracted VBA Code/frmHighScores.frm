VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHighScores 
   Caption         =   "High Scores"
   ClientHeight    =   3840
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6030
   OleObjectBlob   =   "frmHighScores.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==== Snake Game by Mitchell Marino ========================+'
' Name: Mitchell Marino
' Date: 03/04/2017
' Program title: SnakeGame.xlsm
' Description: Highscoreframe functionality.
'===========================================================+'
Option Explicit

'Connection.
Private cn As New ADODB.Connection
'Recordset.
Private rsScores As New ADODB.Recordset

'Ok button click action.
Private Sub btnOk_Click()
    'Exit userform.
    Unload Me
End Sub


Private Sub lblScores_Click()

End Sub


'Initialize the userform.
Private Sub userForm_Initialize()
    
    'Open a connection.
    Call Open_Connection(Session_FilePath)
    
    'Generate a customer list using the private subroutine Generate_Customer_List.
    Call Generate_Score_List

    'Set the list index for customers list to 0.
    lstScores.ListIndex = 0
    
End Sub

'Opens a connection.
Private Sub Open_Connection(filepath As String)

    'Opening the connection.
    With cn
        .ConnectionString = "Data Source=" & filepath
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With

End Sub

'Runs a query to obtain appropriate product record information from database.
Public Sub Generate_Score_Record()
    
    'Temporary string used to run query.
    Dim SQL As String
    
    'Obtain FirstName, and LastName from Scores, then order by score in descending order.
    SQL = "SELECT FirstName, LastName, Score FROM Scores ORDER By Score DESC;"
    rsScores.Open SQL, cn 'Run query.

End Sub

'Generate score list.
Private Sub Generate_Score_List()
    
    Dim i As Integer 'Counter variable.
    i = 0
    
    'Generate a score record.
    Call Generate_Score_Record

    'RS now holds score records.
    With rsScores
        'Until end of record,
        Do Until .EOF
            'Add an item to the list.
            lstScores.AddItem
            'Adjust the item's values accordingly.
            lstScores.List(i, 0) = i + 1
            lstScores.List(i, 1) = .Fields("FirstName") & " " & .Fields("LastName")
            lstScores.List(i, 2) = .Fields("Score")
            i = i + 1 'Increment index counter.
            .MoveNext 'Next item.
        Loop
    End With
    
    rsScores.Close 'Close record set.
    cn.Close       'Close connection.

End Sub
