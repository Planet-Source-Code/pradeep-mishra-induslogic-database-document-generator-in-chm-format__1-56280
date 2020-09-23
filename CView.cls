VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ViewObject As View
Private strHTML As String
Private strTMPHTML As String

Public Property Get ViewObject() As View
    Set ViewObject = m_ViewObject
End Property

Public Property Set ViewObject(ByVal vNewValue As View)
    Set m_ViewObject = vNewValue
End Property

Public Function Header() As String
    strTMPHTML = m_ViewObject.owner & "." & m_ViewObject.Name
    strHTML = StartHead & StartTitle & "View Details: " & strTMPHTML & EndTitle
    strHTML = strHTML & Replace(StartDIV, "<%HEADING%>", "View Details :" & strTMPHTML) & "<BR>"
    Header = strHTML
End Function

Public Function Properties() As String
    strHTML = PropViewHeadHTML
    strHTML = Replace(strHTML, "<%CREATION_DATE%>", m_ViewObject.CreateDate)
    strHTML = Replace(strHTML, "<%IS_SCHEMA_BOUND%>", m_ViewObject.Properties(6).Value)
    Properties = strHTML & "<br>"
End Function

Public Function Column() As String
    Dim objViewColumn As Object
    Dim strColumnHTML As String
    Dim i As Long
    
    strHTML = PropViewColumnTabStartHTML & PropViewColumnHeadHTML
    strColumnHTML = PropViewColumnDetailsHTML
    
    For i = 1 To m_ViewObject.ListColumns.Count
        Set objViewColumn = m_ViewObject.ListColumns(i)
        strColumnHTML = strColumnHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
        With objViewColumn
            strColumnHTML = strColumnHTML & "<TD>" & .Name & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .Datatype & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .length & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .NumericPrecision & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .NumericScale & "</TD>"
            strColumnHTML = strColumnHTML & "</TR>"
        End With
    Next
    strHTML = strHTML & strColumnHTML & PropViewColumnTabEndHTML
    Column = strHTML & "<br>"
End Function

Public Function Source() As String
    Dim strSource As String
    
    strHTML = PropViewSourceHTML
    strSource = m_ViewObject.Script()
    strSource = Replace(strSource, vbCrLf, "<BR>")
    strSource = Replace(strSource, vbTab, "&nbsp;&nbsp;&nbsp")
    strHTML = Replace(strHTML, "<%SOURCE%>", strSource)
    
    Source = strHTML
End Function

