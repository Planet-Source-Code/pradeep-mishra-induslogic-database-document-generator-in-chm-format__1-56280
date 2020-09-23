VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sPath As String
Private m_sKey As String
Private m_sSection As String
Private m_sDefault As String
Private m_lLastReturnCode As Long


'Declares for cIniFile:
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Property Get LastReturnCode() As Long
    'Did the last call succeed?
    '0 if not!
    LastReturnCode = m_lLastReturnCode
End Property

Property Let Default(sDefault As String)
    'What to return if something goes wrong
    '    :
    m_sDefault = sDefault
End Property

Property Get Default() As String
    'What to return if something goes wrong
    '    :
    Default = m_sDefault
End Property

Property Let Path(sPath As String)
    'The filename of the INI file:
    m_sPath = sPath
End Property

Property Get Path() As String
    'The filename of the INI file:
    Path = m_sPath
End Property

Property Let Key(sKey As String)
    'The KEY= bit to look for
    m_sKey = sKey
End Property

Property Get Key() As String
    'The KEY= bit to look for
    Key = m_sKey
End Property

Property Let Section(sSection As String)
    'The [SECTION] bit to look for
    m_sSection = sSection
End Property

Property Get Section() As String
    'The [SECTION] bit to look for
    Section = m_sSection
End Property

'Get the value of the current Key within Section of Path
Property Get Value() As String
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    sBuf = Space(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, iSize, m_sPath)


    If (iSize > 0) Then
        Value = Left(sBuf, iRetCode)
    Else
        Value = ""
    End If
End Property

Property Let Value(sValue As String)
    Dim iPos As Integer
    iPos = InStr(sValue, Chr(0))

    Do While iPos <> 0
        sValue = Left(sValue, (iPos - 1)) & Mid(sValue, (iPos + 1))
        iPos = InStr(sValue, Chr(0))
    Loop
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, sValue, m_sPath)
End Property

'Delete the value at Key within Sectionof Path
Public Sub DeleteValue()
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, 0&, m_sPath)
End Sub

'Delete the Section in Path
Public Sub DeleteSection()
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, 0&, m_sPath)
End Sub

'Return all the keys and values within the current section, separated by chr(0):
Property Get INISection() As String
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    sBuf = Space$(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(m_sSection, 0&, m_sDefault, sBuf, iSize, m_sPath)

    If (iSize > 0) Then
        INISection = Left$(sBuf, iRetCode)
    Else
        INISection = ""
    End If
End Property

'Set one or more the keys within the current section. Keys and Values should be separated by chr$(0):
Property Let INISection(sSection As String)
    m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, sSection, m_sPath)
End Property

