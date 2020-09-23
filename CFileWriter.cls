VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FileName As String
Private m_FileData As String
Private m_Path As String

Public Property Get filename() As String
    filename = m_FileName
End Property

Public Property Let filename(ByVal vNewValue As String)
    m_FileName = vNewValue
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal vNewValue As String)
    m_Path = vNewValue
End Property

Public Property Get FileData() As String
    FileData = m_FileData
End Property

Public Property Let FileData(ByVal vNewValue As String)
    m_FileData = vNewValue
End Property

Public Function WriteToFile()
    Dim oFSO As Scripting.FileSystemObject
    Dim oFileWriter As Scripting.TextStream
    Dim vFreeFile As Long
    
    Set oFSO = New Scripting.FileSystemObject
    
    
    If Not oFSO.FolderExists(m_Path) Then
        oFSO.CreateFolder m_Path
    End If
    
    vFreeFile = FreeFile
    Open m_Path & "\" & m_FileName For Output As #vFreeFile
        Print #vFreeFile, m_FileData
    Close #vFreeFile
    
    sHHFiles = sHHFiles & m_FileName & vbCrLf
    Set oFSO = Nothing
End Function

Public Function ReadFromFile() As String
    Dim vFreeFile As Long
    Dim str As String
    
    vFreeFile = FreeFile
    Open m_Path & "\" & m_FileName For Input As #vFreeFile
        str = Input(LOF(vFreeFile), 1)
    Close #vFreeFile
    ReadFromFile = str
End Function

