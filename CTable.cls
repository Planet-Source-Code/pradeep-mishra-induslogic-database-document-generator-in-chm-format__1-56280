VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TableObject As Table
Private strHTML As String
Private strTMPHTML As String

Public Property Get TableObject() As Table
    Set TableObject = m_TableObject
End Property

Public Property Set TableObject(ByVal vNewValue As Table)
    Set m_TableObject = vNewValue
End Property

Public Function Header() As String
    strTMPHTML = m_TableObject.owner & "." & m_TableObject.Name
    strHTML = StartHead & StartTitle & "Table Details: " & strTMPHTML & EndTitle
    strHTML = strHTML & Replace(StartDIV, "<%HEADING%>", "Table Details :" & strTMPHTML) & "<BR>"
    Header = strHTML
End Function

Public Function Properties() As String
    strHTML = PropTableHeadHTML
    strHTML = Replace(strHTML, "<%CREATION_DATE%>", m_TableObject.CreateDate)
    strHTML = Replace(strHTML, "<%DATA_SPACE_USED%>", m_TableObject.DataSpaceUsed)
    strHTML = Replace(strHTML, "<%NO_OF_ROWS%>", m_TableObject.rows)
    Properties = strHTML & "<br>"
End Function

Public Function Documentation() As String
    Dim strDesc As String
    
    strDesc = GetTableDescription(m_TableObject)
    
    If strDesc <> "" Then
        strHTML = PropTableDescHTML
        strHTML = Replace(strHTML, "<%DESCRIPTION%>", strDesc)
        Documentation = strHTML & "<br>"
    Else
        Documentation = ""
    End If

End Function

Public Function ReferencedTables() As String
    Dim objParam As QueryResults
    Dim strColumnHTML As String
    
    Dim strTable As String
    Dim strKey As String
    Dim strReferenced As String
    Dim i As Long
    
    strHTML = ""
    
    Set objParam = m_TableObject.EnumReferencedKeys
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                strHTML = PropTableReferencedTablesHTML
                For i = 1 To .rows
                    strTable = .GetColumnString(i, 1)
                    strKey = .GetColumnString(i, 2)
                    strReferenced = IIf(.GetColumnString(i, 3) = True, "True", "False")
                    
                    strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
                    strHTML = strHTML & "<TD><A HREF='" & "Table." & Replace(Replace(strTable, "[", ""), "]", "") & ".htm'>" & strTable & "</A></TD>"
                    strHTML = strHTML & "<TD>" & strKey & "</TD>"
                    strHTML = strHTML & "<TD>" & strReferenced & "</TD>"
                    strHTML = strHTML & "</TR>"
                Next
                strHTML = strHTML & "</TBODY>" & "</TABLE>" & "<BR>"
            End If
        End If
    End With
    
    ReferencedTables = strHTML
End Function

Public Function ReferencingTables() As String
    Dim objParam As QueryResults
    Dim strColumnHTML As String
    
    Dim strTable As String
    Dim strKey As String
    Dim strReferenced As String
    Dim i As Long
    
    strHTML = ""
    
    Set objParam = m_TableObject.EnumReferencingKeys()
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                strHTML = PropTableReferencingTablesHTML
                For i = 1 To .rows
                    strTable = .GetColumnString(i, 1)
                    strKey = .GetColumnString(i, 2)
                    strReferenced = IIf(.GetColumnString(i, 3) = True, "True", "False")
                    
                    strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
                    strHTML = strHTML & "<TD><A HREF='" & "Table." & Replace(Replace(strTable, "[", ""), "]", "") & ".htm'>" & strTable & "</A></TD>"
                    strHTML = strHTML & "<TD>" & strKey & "</TD>"
                    strHTML = strHTML & "<TD>" & strReferenced & "</TD>"
                    strHTML = strHTML & "</TR>"
                Next
                strHTML = strHTML & "</TBODY>" & "</TABLE>" & "<BR>"
            End If
        End If
    End With
    
    ReferencingTables = strHTML
End Function

Public Function Indexes() As String
    Dim ctrIndex As Long
    Dim i As Long
    
    strHTML = ""
    
    If m_TableObject.Indexes.Count > 0 Then
        strHTML = PropTableIndexStartHTML
        strHTML = strHTML & PropTableIndexHeadHTML & PropTableIndexDetailsHTML
        
        For ctrIndex = 1 To m_TableObject.Indexes.Count
            strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
            With m_TableObject.Indexes(ctrIndex)
                'If Left(.Name, 1) <> "_" Then
                strHTML = strHTML & "<TD>" & .Name & "</TD>"
                strHTML = strHTML & "<TD>" & ListIndexedColumns(m_TableObject.Indexes(ctrIndex)) & "</TD>"
                strHTML = strHTML & "<TD>" & IIf((.Name = PrimaryKey(m_TableObject)), "True", "False") & "</TD>"
                strHTML = strHTML & "<TD>" & IIf((.Name = ClusteredIndex(m_TableObject)), "True", "False") & "</TD>"
                strHTML = strHTML & "<TD>" & IIf((InStr(m_TableObject.Indexes(ctrIndex).Script(), "UNIQUE") > 0), "True", "False") & "</TD>"
                strHTML = strHTML & "<TD>" & .SpaceUsed & "</TD>"
                strHTML = strHTML & "</TR>"
                i = i + 1
                'End If
            End With
        Next
        
        strHTML = strHTML & PropTableIndexDetailsHTML & PropTableIndexEndHTML
        strHTML = strHTML & "<br>"
    End If
    Indexes = strHTML
End Function

Public Function Column() As String
    Dim objTabColumn As Column
    Dim strColumnHTML As String
    Dim i As Long
    
    strHTML = PropTableColumnTabStartHTML & PropTableColumnHeadHTML
    
    strColumnHTML = PropTableColumnDetailsHTML
    For Each objTabColumn In m_TableObject.Columns
        strColumnHTML = strColumnHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
        With objTabColumn
            strColumnHTML = strColumnHTML & "<TD>" & .Name & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & GetColumnDescription(m_TableObject, .Name) & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .InPrimaryKey & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .dataType & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .length & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .NumericPrecision & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .NumericScale & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .AllowNulls & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .Default & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .Rule & "</TD>"
            strColumnHTML = strColumnHTML & "<TD>" & .Identity & "</TD>"
            strColumnHTML = strColumnHTML & "</TR>"
            i = i + 1
        End With
    Next
    strHTML = strHTML & strColumnHTML & PropTableColumnTabEndHTML
    Column = strHTML & "<br>"
End Function

Public Function Dependant() As String
    Dim objParam As QueryResults
    Dim strColumnHTML As String
    
    Dim strName As String
    Dim strType As String
    Dim strOwner As String
    Dim i As Long
    
    strHTML = ""
    
    Set objParam = m_TableObject.EnumDependencies()
    With objParam
        If .ResultSets > 0 Then
            If .rows > 0 Then
                strHTML = PropTableDependantHTML
                For i = 1 To .rows
                    strName = .GetColumnString(i, 2)
                    strType = GetObjDesc(.GetColumnString(i, 1))
                    strOwner = .GetColumnString(i, 3)
                    
                    strHTML = strHTML & "<TR class=RowColour_" & IIf((i Mod 2 = 0), 1, 2) & ">"
                    strHTML = strHTML & "<TD><A HREF='" & ReturnFileName(strType, strOwner, strName) & "'>" & strName & "</A></TD>"
                    strHTML = strHTML & "<TD>" & strType & "</TD>"
                    strHTML = strHTML & "<TD>" & .GetColumnString(i, 4) & "</TD>"
                    strHTML = strHTML & "</TR>"
                Next
                strHTML = strHTML & "</TBODY>" & "</TABLE>" & "<BR>"
            End If
        End If
    End With
    
    Dependant = strHTML
End Function
