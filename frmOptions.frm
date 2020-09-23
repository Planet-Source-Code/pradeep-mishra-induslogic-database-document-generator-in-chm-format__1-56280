VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTML Compiler Options"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleMode       =   0  'User
   ScaleWidth      =   7412.151
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   960
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Document following database objects"
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   " Process Removal Options"
      Top             =   2160
      Width           =   7005
      Begin VB.CheckBox chkDatabase 
         Caption         =   "Database Details"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkTables 
         Caption         =   "Tables"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   2535
      End
      Begin VB.CheckBox chkViews 
         Caption         =   "Views"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   585
         Width           =   2535
      End
      Begin VB.CheckBox chkFunctions 
         Caption         =   "Functions"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   270
         Width           =   2655
      End
      Begin VB.CheckBox chkStoredProcedures 
         Caption         =   "Stored Procedures"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   915
         Width           =   3075
      End
      Begin VB.CheckBox chkTriggers 
         Caption         =   "Triggers"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   585
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5985
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame fraFrame 
      Height          =   1875
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Default File Path Options"
      Top             =   0
      Width           =   6960
      Begin VB.CheckBox chkDeleteFiles 
         Caption         =   "Delete temporary files after generating compiled html help"
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   1380
         Width           =   4935
      End
      Begin VB.CheckBox chkShowHTMLCompiler 
         Caption         =   "Show html compiler DOS window."
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse..."
         Height          =   315
         Left            =   5700
         TabIndex        =   3
         Top             =   540
         Width           =   915
      End
      Begin VB.TextBox txtDestDir 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Text            =   "C:\Program Files\HTML Help Workshop\hhc.exe"
         Top             =   540
         Width           =   5535
      End
      Begin VB.Label lblLabel 
         Caption         =   "HTML help compiler path"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmOptions
' DateTime  : 9/6/2004 12:49
' Author    : pradeep.mishra
' Purpose   : Provide Options for Documentation
'---------------------------------------------------------------------------------------
Option Explicit

Dim cIni As CIniFile

Private Sub cmdApply_Click()
    SaveSettings
End Sub

'------------------------------------------------------------------------------------
'Function       : cmdBrowse_Click
'Description    : Called when user presses browse button
'Parameters     :
'Creator        : pradeep.mishra
Private Sub cmdBrowse_Click()
    Dim getDir As String
    
    With dlgCommon
        .InitDir = "C:\Program Files\HTML Help Workshop"
        .Filter = "Html Help Compiler(HHC.EXE)|HHC.EXE"
        .ShowOpen
        .CancelError = False
        txtDestDir.Text = IIf((Len(.filename) > 0), .filename, txtDestDir.Text)
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSettings
    Unload Me
End Sub

Private Sub Form_Load()
    Set cIni = New CIniFile
    LoadSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettings
    Set cIni = Nothing

End Sub

'------------------------------------------------------------------------------------
'Function       : SaveSettings
'Description    : Saves Settings to IniFile
'Parameters     :
'------------------------------------------------------------------------------------
Sub SaveSettings()
    With cIni
        .Path = "IDBDocumentor.INI"
        
        .Section = "HTMLHelpOptions"
        
        .Key = "ShowHTMLCompiler"
        .Value = chkShowHTMLCompiler
        
        .Key = "LeaveFiles"
         .Value = chkDeleteFiles
         
         .Key = "HTMLHELPCompilerPath"
         .Value = txtDestDir.Text
         
         .Section = "DatabaseOptions"
         
         .Key = "Tables"
         .Value = chkTables
         
         .Key = "Views"
         .Value = chkViews
         
         .Key = "StoredProcedures"
         .Value = chkStoredProcedures
         
         .Key = "Functions"
         .Value = chkFunctions

         .Key = "Triggers"
         .Value = chkTriggers
         
         .Key = "Database"
         .Value = chkDatabase
         
    End With
End Sub

'------------------------------------------------------------------------------------
'Function       : LoadSettings
'Description    : Loads Settings from INI FIle
'Parameters     :
'------------------------------------------------------------------------------------
Sub LoadSettings()
    With cIni
        .Path = "IDBDocumentor.INI"
        .Section = "HTMLHelpOptions"
        
        .Key = "ShowHTMLCompiler"
        chkShowHTMLCompiler = Val(.Value)
        
        .Key = "LeaveFiles"
         chkDeleteFiles = Val(.Value)
        
         .Key = "HTMLHELPCompilerPath"
         txtDestDir = .Value
         
         .Section = "DatabaseOptions"
         
         .Key = "Tables"
         chkTables = Val(.Value)
         
         .Key = "Views"
         chkViews = Val(.Value)
         
         .Key = "StoredProcedures"
         chkStoredProcedures = Val(.Value)
         
         .Key = "Functions"
         chkFunctions = Val(.Value)

         .Key = "Triggers"
         chkTriggers = Val(.Value)
         
         .Key = "Database"
         chkDatabase = Val(.Value)
        
    End With
End Sub
