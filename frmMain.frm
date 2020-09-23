VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Server Documentation Builder"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser wbMain 
      Height          =   1635
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   8115
      ExtentX         =   14314
      ExtentY         =   2884
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   60
      TabIndex        =   2
      Top             =   6300
      Width           =   8115
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   6540
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   6540
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   6540
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      Top             =   6540
      Width           =   1275
   End
   Begin VB.Frame fmaMain 
      Height          =   3600
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   8115
   End
   Begin VB.Image imgEngeniaImage 
      Height          =   960
      Left            =   60
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oHDoc As HTMLDocument
Dim CurrentForm As String

Private Sub cmdBack_Click()
    Select Case CurrentForm
        Case "SQLServer"
            SetParent frmBlank.hWnd, fmaMain.hWnd
            frmBlank.Show
            frmBlank.Move 120, 120
            CurrentForm = "Main"
            oHDoc.body.innerHTML = HHeader1
            cmdBack.Enabled = False
        Case "SelectDirectory"
            SetParent frmSQLServer.hWnd, fmaMain.hWnd
            frmSQLServer.Show
            frmSQLServer.Move 120, 120
            CurrentForm = "SQLServer"
            oHDoc.body.innerHTML = HHeader2
        
        Case "StartProcessing"
            SetParent frmSelectDirectory.hWnd, fmaMain.hWnd
            frmSelectDirectory.Show
            frmSelectDirectory.Move 120, 120
            CurrentForm = "SelectDirectory"
            oHDoc.body.innerHTML = HHeader3
            cmdNext.Caption = "&Next >"

        Case Else
    End Select
    
    If CurrentForm = "DBObjects" Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdNext_Click()
    Dim frm As Form
    
    If cmdNext.Caption = "&Build Now.." Then
        frmStartProcessing.StartProcessing
        cmdNext.Caption = "&Close"
        Exit Sub
    End If
    
    If cmdNext.Caption = "&Close" Then
        For Each frm In Forms
            Unload frm
        Next
        End
    End If

    Select Case CurrentForm
        Case "Main"
            SetParent frmSQLServer.hWnd, fmaMain.hWnd
            frmSQLServer.Show
            frmSQLServer.Move 120, 120
            CurrentForm = "SQLServer"
            oHDoc.body.innerHTML = HHeader2
            cmdBack.Enabled = True
        
        Case "SQLServer"
            If Not mIsConnected Then
                MsgBox "Please connect to a server and select default database first...", vbCritical, "Connect to Database first"
                Exit Sub
            End If
            
            If Not mIsConnected Or frmSQLServer.cmbServer.Text = "" Or frmSQLServer.cmbDatabase.Text = "" Then
                MsgBox "Please connect to a server and select default database first...", vbCritical, "Connect to Database first"
                Exit Sub
            End If
            
            SetParent frmSelectDirectory.hWnd, fmaMain.hWnd
            frmSelectDirectory.Show
            frmSelectDirectory.Move 120, 120
            CurrentForm = "SelectDirectory"
            oHDoc.body.innerHTML = HHeader3
            thisDB = frmSQLServer.cmbDatabase.Text
        
        Case "SelectDirectory"
            SetParent frmStartProcessing.hWnd, fmaMain.hWnd
            frmStartProcessing.Show
            frmStartProcessing.Move 120, 120
            CurrentForm = "StartProcessing"
            oHDoc.body.innerHTML = HHeader4
            
            cmdNext.Enabled = True
            cmdNext.Caption = "&Build Now.."
        Case Else
    
    End Select
    
    If CurrentForm = "Main" Then
        cmdBack.Enabled = False
    Else
        cmdBack.Enabled = True
    End If

End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub Form_Load()
    CurrentForm = "Main"
    wbMain.navigate "ABOUT:BLANK"
    Set oHDoc = wbMain.document
    
    Do While wbMain.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    
    oHDoc.body.innerHTML = HHeader1
    
    If CurrentForm = "Main" Then cmdBack.Enabled = False
End Sub

'Make sure that you unload all the forms when main window is closed
Private Sub Form_Unload(Cancel As Integer)
    Dim fObj As Form
    
    For Each fObj In Forms
        Unload fObj
    Next
End Sub

