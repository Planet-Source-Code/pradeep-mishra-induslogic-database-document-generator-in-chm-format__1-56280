VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartProcessing 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMessages 
      Height          =   1935
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1020
      Width           =   7635
   End
   Begin MSComctlLib.ProgressBar pbCurrentObject 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   90
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ProgressBar pbOverAll 
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   420
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblLabel 
      Caption         =   "Output Messages"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   1755
   End
   Begin VB.Label lblLabel 
      Caption         =   "Overall Progress"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Caption         =   "Object Progress"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmStartProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cIni As CIniFile

Dim txtDestDir As String
Dim chkShowHTMLCompiler As Long
Dim chkDeleteFiles As Long
Dim chkTables As Long
Dim chkViews As Long
Dim chkStoredProcedures As Long
Dim chkFunctions As Long
Dim chkTriggers As Long
Dim chkDatabase As Long

Public Function StartProcessing()
    On Error GoTo errorHandler
    
    Dim objDoc As CDocumentGenerator
    Dim objFileWriter As CFileWriter
    Dim strTemp As String
    Dim filename As String
    
    Set oDatabase = mServer.Databases(thisDB)
    
    txtMessages.Text = "Generating HTML documentation..." & vbCrLf
    
    strTemp = Replace(TOCDatabase, "<%DATABASE%>", thisDB)
    strTemp = Replace(strTemp, "<%DBDOCUMENT%>", thisDB & ".htm")
    TOC = TOCHeader & Replace(TOCServer, "<%SERVER%>", thisServer) & strTemp & TOCDatabaseD
    
    strIndex = IndexHeader
    
    pbOverAll.Min = 1
    pbOverAll.Max = chkDatabase + chkTables + chkViews + chkFunctions + chkStoredProcedures + chkTriggers + 1
    
    Set objDoc = New CDocumentGenerator
    With objDoc
        If chkDatabase Then
            txtMessages.Text = txtMessages.Text & "Generating database specific document..." & vbCrLf
            .DatabaseDetail
            pbOverAll.Value = 1
        End If
        
        If chkTables Then
            txtMessages.Text = txtMessages.Text & "Generating table documentation..." & vbCrLf
            TOC = TOC & Replace(TOCTableH, "<%TABLELINK%>", "Table." & MakeCompatibleFileName(oDatabase.owner) & ".AllTables" & ".htm")
            .TableDOC
            TOC = TOC & "</UL>"
            pbOverAll.Value = pbOverAll.Value + 1
        End If
        
        If chkTriggers Then
            txtMessages.Text = txtMessages.Text & "Generating trigger documentation..." & vbCrLf
            TOC = TOC & TOCTriggerH
            .TriggerDOC
            TOC = TOC & "</UL>"
            pbOverAll.Value = pbOverAll.Value + 1
        End If
        
        If chkViews Then
            txtMessages.Text = txtMessages.Text & "Generating Views documentation..." & vbCrLf
            TOC = TOC & TOCViewH
            .ViewsDOC
            TOC = TOC & "</UL>"
            pbOverAll.Value = pbOverAll.Value + 1
        End If
        
        If chkStoredProcedures Then
            txtMessages.Text = txtMessages.Text & "Generating Stored Procedure documentation..." & vbCrLf
            TOC = TOC & TOCStoredProcedureH
            .StoredProcedureDOC
            TOC = TOC & "</UL>"
            pbOverAll.Value = pbOverAll.Value + 1
        End If
        
        If chkFunctions Then
            txtMessages.Text = txtMessages.Text & "Generating User Defined Functions documentation..." & vbCrLf
            TOC = TOC & TOCUserDefinedFunctionH
            .UserDefinedFunctionsDOC
            TOC = TOC & "</UL>"
            pbOverAll.Value = pbOverAll.Value + 1
        End If
        
        'txtMessages.Text = txtMessages.Text & "Generating Index documentation..." & vbCrLf
        '.IndexesDOC
        'pbOverAll.Value = pbOverAll.Value + 1
        
    End With
    
    txtMessages.Text = txtMessages.Text & "Compiling HTML Help files..." & vbCrLf
    
    Set objFileWriter = New CFileWriter
    With objFileWriter
        .filename = "TOC.hhc"
        .Path = frmSelectDirectory.txtDirectory
        TOC = Replace(TOC, "'", Chr(34)) & "</UL></UL></BODY></HTML>"
        strIndex = Replace(strIndex, "'", Chr(34)) & "</UL></BODY></HTML>"
        .FileData = TOC
        .WriteToFile
        
        .FileData = strIndex
        .filename = "Index.hhk"
        .WriteToFile
        
        sHHHead = Replace(HHHead, "<%COMPILEDFILE%>", thisServer & "." & thisDB & ".chm")
        sHHHead = Replace(sHHHead, "<%TITLE%>", thisServer & "." & thisDB)
        sHHFiles = HHFiles & sHHFiles
        HHFileName = thisServer & "." & thisDB & "." & "hhp"
        .filename = HHFileName
        .Path = frmSelectDirectory.txtDirectory
        .FileData = sHHHead & sHHFiles
        .WriteToFile
        
        .filename = "Style.CSS"
        .Path = App.Path
        strTemp = .ReadFromFile
        
        .filename = "Style.CSS"
        .Path = frmSelectDirectory.txtDirectory
        .FileData = strTemp
        .WriteToFile
        
    End With
    Set objFileWriter = Nothing
    
    Dim idProg, iExit As Long
    Dim strShell As String
    
    strShell = Chr(34) & txtDestDir & Chr(34) & " " & Chr(34) & frmSelectDirectory.txtDirectory & "\" & HHFileName & Chr(34)
    
    If chkShowHTMLCompiler = 0 Then
        idProg = Shell(strShell, vbHide)
    Else
        idProg = Shell(strShell, vbNormalFocus)
    End If
    
    iExit = fWait(idProg)
    
    If chkDeleteFiles Then
        On Error Resume Next
        txtMessages.Text = txtMessages.Text & "Deleting temporary files..." & vbCrLf
        Kill frmSelectDirectory.txtDirectory & "\*.htm"
        Kill frmSelectDirectory.txtDirectory & "\*.Css"
        Kill frmSelectDirectory.txtDirectory & "\*.HHC"
        Kill frmSelectDirectory.txtDirectory & "\*.HHP"
        Kill frmSelectDirectory.txtDirectory & "\*.HHK"
        
        txtMessages.Text = txtMessages.Text & "Finished Processing..." & vbCrLf & _
            "In case you want to compile the files your self the command line for that is " & strShell & vbCrLf
        
        On Error GoTo errorHandler
    End If
    
    pbOverAll.Value = pbOverAll.Value + 1
    
    If idProg = 0 Then
        Err.Raise 10001, "Startprocessing", "There was an error while running Microsoft HTML help workshop"
    Else
        txtMessages.Text = txtMessages.Text & "Successfully created database document, please check " & HHFileName & " in " & frmSelectDirectory.txtDirectory & " directory." & vbCrLf
        MsgBox "Successfully created database document, please check " & HHFileName & " in " & frmSelectDirectory.txtDirectory & " directory.", vbInformation
    End If
    
    Exit Function
errorHandler:
    Dim Choice As Long
    Dim strMessageString As String
    Select Case Err.Number
        Case 76, 53
            strMessageString = "HTML Help Workshop Not found, please make sure that file " & HTMLCompiler & " exists."
        Case Else
            strMessageString = "An Error " & Err.Description & " occured."
    End Select

    Choice = MsgBox(strMessageString & " What do you want to do..", vbCritical + vbAbortRetryIgnore, "Errors occured...")
    Select Case Choice
        Case vbAbort
            Exit Function
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Function

Function fWait(ByVal lProgID As Long) As Long
    ' Wait until proggie exit code <>
    ' STILL_ACTIVE&
    Dim lExitCode As Long, hdlProg As Long
    ' Get proggie handle
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    ' Get current proggie exit code
    GetExitCodeProcess hdlProg, lExitCode


    Do While lExitCode = STILL_ACTIVE
        DoEvents
        GetExitCodeProcess hdlProg, lExitCode
    Loop
    fWait = lExitCode
    
    CloseHandle hdlProg
End Function

Private Sub Form_Load()
    Set cIni = New CIniFile
    LoadSettings
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

