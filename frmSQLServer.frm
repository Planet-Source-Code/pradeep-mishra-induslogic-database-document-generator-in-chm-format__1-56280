VERSION 5.00
Begin VB.Form frmSQLServer 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDatabase 
      Height          =   315
      Left            =   6060
      Picture         =   "frmSQLServer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2460
      Width           =   495
   End
   Begin VB.CommandButton cmdServer 
      Height          =   315
      Left            =   6060
      Picture         =   "frmSQLServer.frx":02C2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   495
   End
   Begin VB.ComboBox cmbDatabase 
      Height          =   315
      ItemData        =   "frmSQLServer.frx":0584
      Left            =   1140
      List            =   "frmSQLServer.frx":0586
      TabIndex        =   10
      Top             =   2460
      Width           =   4875
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   8
      Text            =   "sa"
      Top             =   1890
      Width           =   4815
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Text            =   "sa"
      Top             =   1560
      Width           =   4815
   End
   Begin VB.OptionButton optOption 
      Caption         =   "Use SQL Server Authentication"
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   5775
   End
   Begin VB.OptionButton optOption 
      Caption         =   "Use Windows Authentication"
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   660
      Value           =   -1  'True
      Width           =   5775
   End
   Begin VB.ComboBox cmbServer 
      Height          =   315
      ItemData        =   "frmSQLServer.frx":0588
      Left            =   1140
      List            =   "frmSQLServer.frx":058A
      TabIndex        =   1
      Top             =   180
      Width           =   4875
   End
   Begin VB.Label lblLabel 
      Caption         =   "Database"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   2490
      Width           =   915
   End
   Begin VB.Label lblLabel 
      Caption         =   "Password"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   7
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label lblLabel 
      Caption         =   "User ID"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblLabel 
      Caption         =   "Server"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   915
   End
End
Attribute VB_Name = "frmSQLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbDatabase_DropDown()
    On Error GoTo errorHandler
    Dim strErrorMessage As String
    If Not ConnectToServer(cmbServer.Text, txtUserID, txtPassword, optOption(0).Value, strErrorMessage) Then
        MsgBox strErrorMessage, vbCritical, "Login failed..."
        Exit Sub
    End If
    GetAllDatabases
    Exit Sub

errorHandler:

End Sub

Private Sub cmbServer_Change()
    cmbDatabase.Clear
    mIsConnected = False
End Sub

Private Sub cmbServer_Click()
    cmbDatabase.Clear
    mIsConnected = False
End Sub

Private Sub cmdDatabase_Click()
    cmbDatabase_DropDown
End Sub

Private Sub cmdServer_Click()
    ListAllServers cmbServer
End Sub

Private Sub Form_Load()
    ListAllServers cmbServer
    cmbServer.ListIndex = 0
    optOption_Click 0
End Sub

Private Sub optOption_Click(index As Integer)
    Select Case index
        Case 0
            txtUserID.Enabled = False
            txtUserID.BackColor = &H8000000F
            
            txtPassword.Enabled = False
            txtPassword.BackColor = &H8000000F
            
        Case 1
            txtUserID.Enabled = True
            txtUserID.BackColor = &H80000005
            
            txtPassword.Enabled = True
            txtPassword.BackColor = &H80000005
    End Select
End Sub

Private Function GetAllDatabases()
    Dim objDB As SQLDMO.Database
    
    If cmbDatabase.ListCount = 0 Then
        DoEvents
        Screen.MousePointer = vbHourglass
        
        For Each objDB In mServer.Databases
            cmbDatabase.AddItem objDB.Name
        Next
        Screen.MousePointer = vbDefault
    End If

End Function
