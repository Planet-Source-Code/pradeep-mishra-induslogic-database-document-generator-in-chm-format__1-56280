VERSION 5.00
Begin VB.Form frmSelectDirectory 
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
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txtDirectory 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   90
      Width           =   2835
   End
   Begin VB.Label lblLabel 
      Caption         =   "Directory:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSelectDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    Dim strDirectory As String
    strDirectory = BrowseForFolder(Me, "Select a directory where you want SQL Server documentation files to be generated", txtDirectory.Text)
    
    If strDirectory <> "" Then
        txtDirectory.Text = strDirectory
    End If
End Sub

Private Sub cmdBrowse_LostFocus()
    frmMain.SetFocus
End Sub

Private Sub Form_Load()
    txtDirectory.Text = GetSetting(App.EXEName, "Settings", "DefaultSaveDirectory", "C:\")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Settings", "DefaultSaveDirectory", txtDirectory.Text
End Sub
