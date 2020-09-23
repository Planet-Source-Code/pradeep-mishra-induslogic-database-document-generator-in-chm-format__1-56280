VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScriptDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 '************************************************************************
'  Script a Sql Server
'
'  Parameters (DTSGlobalVariables)
'    ServerName          - Net Name of SQL Server being checked in
'    WorkingFolderRoot   - local file system folder to store scripts
'
'  See
'  \Program Files\Microsoft SQL Server\80\Tools\DevTools\Include\sqldmo.h
'  for interesting comments concerning usage of the enums that have been
'  explicitly declared as constants in this script.
'
'  Script methods of Table and BackupDevice require additional parameter
'
'  Bill Wunder
'************************************************************************

Option Explicit

' Global vars

    Dim oSQLServer, oDatabase
    Dim oFileSystem, oFile, sFileName
    Dim sWorkingFolderRoot, sWorkingFolder
    Dim iScriptType, iScript2Type
    Dim sTextStream

'Constants
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8

    Const TristateUseDefault = -2
    Const TristateTrue = -1
    Const TristateFalse = 0

    'enum SQLDMO.SQLDMO_SCRIPT_TYPE
    Const SQLDMOScript_UseQuotedIdentifiers = -1
    Const SQLDMOScript_None = 0
    Const SQLDMOScript_Drops = 1
    Const SQLDMOScript_ObjectPermissions = 2
    Const SQLDMOScript_Default = 4
    Const SQLDMOScript_PrimaryObject = 4
    Const SQLDMOScript_ClusteredIndexes = 8
    Const SQLDMOScript_Triggers = 16
    Const SQLDMOScript_DatabasePermissions = 32
    Const SQLDMOScript_Permissions = 34
    Const SQLDMOScript_ToFileOnly = 64
    Const SQLDMOScript_Bindings = 128
    Const SQLDMOScript_AppendToFile = 256
    Const SQLDMOScript_NoDRI = 512
    Const SQLDMOScript_UDDTsToBaseType = 1024
    Const SQLDMOScript_IncludeIfNotExists = 4096
    Const SQLDMOScript_NonClusteredIndexes = 8192
    Const SQLDMOScript_Aliases = 16384
    Const SQLDMOScript_NoCommandTerm = 32768
    Const SQLDMOScript_DRIIndexes = 65536
    Const SQLDMOScript_Indexes = 73736
    Const SQLDMOScript_IncludeHeaders = 131072
    Const SQLDMOScript_OwnerQualify = 262144
    Const SQLDMOScript_TransferDefault = 422143
    Const SQLDMOScript_TimestampToBinary = 524288
    Const SQLDMOScript_SortedData = 1048576
    Const SQLDMOScript_SortedDataReorg = 2097152
    Const SQLDMOScript_DRI_NonClustered = 4194304
    Const SQLDMOScript_DRI_Clustered = 8388608
    Const SQLDMOScript_DRI_Checks = 16777216
    Const SQLDMOScript_DRI_Defaults = 33554432
    Const SQLDMOScript_DRI_UniqueKeys = 67108864
    Const SQLDMOScript_DRI_ForeignKeys = 134217728
    Const SQLDMOScript_DRI_PrimaryKey = 268435456
    Const SQLDMOScript_DRI_AllKeys = 469762048
    Const SQLDMOScript_DRI_AllConstraints = 520093696
    Const SQLDMOScript_DRIWithNoCheck = 536870912
    Const SQLDMOScript_DRI_All = 532676608
    Const SQLDMOScript_NoIdentity = 1073741824

    'enum SQLDMO.SQLDMO_SCRIPT2_TYPE
    Const SQLDMOScript2_Default = 0
    Const SQLDMOScript2_AnsiPadding = 1
    Const SQLDMOScript2_AnsiFile = 2
    Const SQLDMOScript2_UnicodeFile = 4
    Const SQLDMOScript2_NonStop = 8
    Const SQLDMOScript2_NoFG = 16
    Const SQLDMOScript2_MarkTriggers = 32
    Const SQLDMOScript2_OnlyUserTriggers = 64
    Const SQLDMOScript2_EncryptPWD = 128
    Const SQLDMOScript2_SeparateXPs = 256
    Const SQLDMOScript2_NoWhatIfIndexes = 512
    Const SQLDMOScript2_AgentNotify = 1024
    Const SQLDMOScript2_AgentAlertJob = 2048
    Const SQLDMOScript2_FullTextIndex = 524288
    Const SQLDMOScript2_LoginSID = 1048576 'BOL is wrong, says 8192
    Const SQLDMOScript2_FullTextCat = 2097152
    Const SQLDMOScript2_ExtendedProperty = 4194304
    Const SQLDMOScript2_NoCollation = 8388608
    Const SQLDMOScript2_JobDisable = 33554432
    Const SQLDMOScript2_ExtendedOnly = 67108864
    Const SQLDMOScript2_70Only = 16777216
    Const SQLDMOScript2_DontScriptJobServer = 134217728

    'enum SQLDMO.SQLDMO_PRIVILEGE_TYPE
    Const SQLDMOPriv_Unknown = 0
    Const SQLDMOPriv_Select = 1
    Const SQLDMOPriv_Insert = 2
    Const SQLDMOPriv_Update = 4
    Const SQLDMOPriv_Delete = 8
    Const SQLDMOPriv_Execute = 16
    Const SQLDMOPriv_References = 32
    Const SQLDMOPriv_AllObjectPrivs = 63 'Default
    Const SQLDMOPriv_CreateTable = 128
    Const SQLDMOPriv_CreateDatabase = 256
    Const SQLDMOPriv_CreateView = 512
    Const SQLDMOPriv_CreateProcedure = 1024
    Const SQLDMOPriv_DumpDatabase = 2048
    Const SQLDMOPriv_CreateDefault = 4096
    Const SQLDMOPriv_DumpTransaction = 8192
    Const SQLDMOPriv_CreateRule = 16384
    Const SQLDMOPriv_DumpTable = 32768
    Const SQLDMOPriv_CreateFunction = 65366
    Const SQLDMOPriv_AllDatabasePrivs = 130944 'Default

Function Main()
  
  ' need a file system instance and a DMO instance
  Set oFileSystem = CreateObject("Scripting.FileSystemObject")
  Set oSQLServer = CreateObject("SQLDMO.SQLServer")
  
  oSQLServer.LoginSecure = True 'trusted
  oSQLServer.Connect DTSGlobalVariables("ServerName").Value
  
  sWorkingFolderRoot = Trim(DTSGlobalVariables("WorkingFolderRoot").Value)

  'verify the path and make sure empty subfolders for the server exists
  SetWorkingFolders
    
  ScriptAlerts
  ScriptBackupDevices
  ScriptJobs
  ScriptLogins
  ScriptOperators

  For Each oDatabase In oSQLServer.Databases
    If Not oDatabase.SystemObject Then

      ScriptDB
      ScriptDataTypes
      ScriptDefaults
      ScriptFunctions
      ScriptProcedures
      ScriptRoles
      ScriptRules
      ScriptTables
      ScriptTriggers
      ScriptUsers
      ScriptViews

    End If
  Next 'oDatabase
    
  oSQLServer.DisConnect
  Set oSQLServer = Nothing

  Set oFileSystem = Nothing

  ''''Main = DTSTaskExecResult_Success
End Function


Function SetWorkingFolders()
  Dim sDBFolder
  Dim oDatabase
  Dim sServerName

  If Right(sWorkingFolderRoot, 1) <> "\" Then
    sWorkingFolderRoot = sWorkingFolderRoot & "\"
  End If
  If oFileSystem.DriveExists(oFileSystem.GetDriveName(sWorkingFolderRoot)) Then
    If oFileSystem.FolderExists(sWorkingFolderRoot) Then
      If Len(oSQLServer.Name) < 2 Then
        ' assume no param provided and script local default instance
        sWorkingFolderRoot = sWorkingFolderRoot & oSQLServer.NetName
      Else
        ' use $ instead of \ for named instance in file system just like M$
        sWorkingFolderRoot = sWorkingFolderRoot & Replace(oSQLServer.Name, "\", "$")
      End If
      If oFileSystem.FolderExists(sWorkingFolderRoot) Then
        oFileSystem.DeleteFolder sWorkingFolderRoot, True 'even if read only
        'sWorkingFolderRoot.Delete '(true)'which one works? seems spotty
      End If
      sWorkingFolderRoot = oFileSystem.CreateFolder(sWorkingFolderRoot) & "\"

      oFileSystem.CreateFolder (sWorkingFolderRoot & "\AgentAlerts")
      oFileSystem.CreateFolder (sWorkingFolderRoot & "\AgentOperators")
      oFileSystem.CreateFolder (sWorkingFolderRoot & "\BackupDevices")
      oFileSystem.CreateFolder (sWorkingFolderRoot & "\Jobs")
      oFileSystem.CreateFolder (sWorkingFolderRoot & "\Logins")

      For Each oDatabase In oSQLServer.Databases
        If Not oDatabase.SystemObject Then
          sDBFolder = oFileSystem.CreateFolder(sWorkingFolderRoot & "\" _
                                               & oDatabase.Name) & "\"
          oFileSystem.CreateFolder (sDBFolder & "\DataTypes")
          oFileSystem.CreateFolder (sDBFolder & "\Defaults")
          oFileSystem.CreateFolder (sDBFolder & "\Functions")
          oFileSystem.CreateFolder (sDBFolder & "\Procedures")
          oFileSystem.CreateFolder (sDBFolder & "\Roles")
          oFileSystem.CreateFolder (sDBFolder & "\Rules")
          oFileSystem.CreateFolder (sDBFolder & "\Tables")
          oFileSystem.CreateFolder (sDBFolder & "\Triggers")
          oFileSystem.CreateFolder (sDBFolder & "\Users")
          oFileSystem.CreateFolder (sDBFolder & "\Views")
        End If
      Next 'oDatabase

    End If
  End If

End Function


Function ScriptDB()

  iScriptType = SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default
  sFileName = sWorkingFolderRoot & oDatabase.Name & "\CreateDB_" _
                                 & oDatabase.Name & ".DB"
 
  oFileSystem.CreateTextFile sFileName, True
  Set oFile = oFileSystem.GetFile(sFileName)
  Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
  sTextStream.Write oDatabase.Script(iScriptType, , iScript2Type)
  sTextStream.Close

End Function


Function ScriptDataTypes()

  Dim oUserDefinedDataType

  iScriptType = SQLDMOScript_Drops _
        Or SQLDMOScript_ObjectPermissions _
        Or SQLDMOScript_Default _
        Or SQLDMOScript_Bindings
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\DataTypes\"
  
  For Each oUserDefinedDataType In oDatabase.UserDefinedDatatypes
    sFileName = sWorkingFolder & oUserDefinedDataType.Name & ".UDT"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write "use " & oDatabase.Name
    sTextStream.WriteBlankLines (1)
    sTextStream.Write "GO"
    sTextStream.WriteBlankLines (1)
    sTextStream.Write oUserDefinedDataType.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'oUserDefinedDataType

End Function


Function ScriptDefaults()

  Dim oDefault

  iScriptType = SQLDMOScript_Drops _
        Or SQLDMOScript_ObjectPermissions _
        Or SQLDMOScript_Default _
        Or SQLDMOScript_Bindings
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Defaults\"

  For Each oDefault In oDatabase.defaults
    sFileName = sWorkingFolder & oDefault.Name & ".DFT"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write "use " & oDatabase.Name
    sTextStream.WriteBlankLines (1)
    sTextStream.Write "GO"
    sTextStream.WriteBlankLines (1)
    sTextStream.Write oDefault.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'oDefault

End Function


Function ScriptFunctions()

  Dim oUserDefinedFunction

  iScriptType = SQLDMOScript_Drops _
        Or SQLDMOScript_ObjectPermissions _
        Or SQLDMOScript_OwnerQualify _
        Or SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Functions\"

  For Each oUserDefinedFunction In oDatabase.UserDefinedFunctions
    If Not oUserDefinedFunction.SystemObject Then
      sFileName = sWorkingFolder & oUserDefinedFunction.owner & "." _
                                 & oUserDefinedFunction.Name & ".UDF"
      oFileSystem.CreateTextFile sFileName, True
      Set oFile = oFileSystem.GetFile(sFileName)
      Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
      sTextStream.Write "use " & oDatabase.Name
      sTextStream.WriteBlankLines (1)
      sTextStream.Write "GO"
      sTextStream.WriteBlankLines (1)
      sTextStream.Write oUserDefinedFunction.Script(iScriptType, , iScript2Type)
      sTextStream.Close
    End If
  Next 'oUserDefinedFunction

End Function


Function ScriptProcedures()

  Dim oStoredProcedure

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_ObjectPermissions _
                Or SQLDMOScript_OwnerQualify _
                Or SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Procedures\"

  For Each oStoredProcedure In oDatabase.StoredProcedures
    If Not oStoredProcedure.SystemObject Then
      'if procedure owner is a domain user the create file will (should)
      'blow up with an invalid path message
      sFileName = sWorkingFolder & Replace(oStoredProcedure.owner, "\", "~") & "." _
                                 & oStoredProcedure.Name & ".PRC"
      oFileSystem.CreateTextFile sFileName, True
      Set oFile = oFileSystem.GetFile(sFileName)
      Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
      sTextStream.Write "use " & oDatabase.Name
      sTextStream.WriteBlankLines (1)
      sTextStream.Write "GO"
      sTextStream.WriteBlankLines (1)
      sTextStream.Write oStoredProcedure.Script(iScriptType, , iScript2Type)
      sTextStream.Close
    End If
  Next 'oStoredProcedure

End Function


Function ScriptRoles()

  Dim oRole
  Dim oMembersQueryResult, iRow
  Dim oDBPermissionsSQLObjectList, oObjectPermissionsSQLObjectList, oPermission
  Dim DoIt

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Roles\"

  For Each oRole In oDatabase.DatabaseRoles
    Set oMembersQueryResult = oRole.EnumDatabaseRoleMember()
    If oMembersQueryResult.rows > 0 And Not (oRole.Name = "public") Then
      DoIt = 1
    End If ' has members

    If Not (oRole.IsFixedRole()) Then
      Set oDBPermissionsSQLObjectList = oRole.ListDatabasePermissions(SQLDMOPriv_AllDatabasePrivs)
      If oDBPermissionsSQLObjectList.Count > 0 Then
        DoIt = 1
      End If 'has db permissions
      Set oObjectPermissionsSQLObjectList = oRole.ListObjectPermissions(SQLDMOPriv_AllObjectPrivs)
      If oObjectPermissionsSQLObjectList.Count > 0 Then
        DoIt = 1
      End If 'has object permissions
    End If 'user defined role

     If DoIt = 1 Then

      sFileName = sWorkingFolder & oRole.Name & ".ROL"
      oFileSystem.CreateTextFile sFileName, True
      Set oFile = oFileSystem.GetFile(sFileName)
      Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
      sTextStream.WriteLine "-- Role: " & oRole.Name
      sTextStream.WriteLine "use " & oDatabase.Name
      sTextStream.WriteBlankLines (1)
      sTextStream.Write "GO"
      sTextStream.WriteBlankLines (1)
 
      If Not (oRole.IsFixedRole()) Then
        sTextStream.Write oRole.Script(iScriptType, , iScript2Type)
      End If

      If Not (oRole.Name = "public") Then ' because everybody is in public
        For iRow = 1 To oMembersQueryResult.rows
          sTextStream.WriteLine "exec sp_addrolemember [" & oRole.Name & "], [" _
                                             & oMembersQueryResult.GetColumnString(iRow, 1) & "]"
          sTextStream.WriteBlankLines (1)
          sTextStream.WriteLine "GO"
          sTextStream.WriteBlankLines (1)
        Next 'Role Member
      End If

      If Not (oRole.IsFixedRole()) Then
        For Each oPermission In oDBPermissionsSQLObjectList
          sTextStream.WriteLine "grant [" & oPermission.PrivilegeTypeName & "] to [" _
                                          & oPermission.Grantee & "]"
          sTextStream.WriteBlankLines (1)
        Next 'Database Permission

        For Each oPermission In oObjectPermissionsSQLObjectList
          sTextStream.WriteLine "grant [" & oPermission.PrivilegeTypeName & "] on [" _
                                      & oPermission.ObjectOwner & "].[" _
                                      & oPermission.ObjectName & "] to [" _
                                      & oPermission.Grantee & "]"
          sTextStream.WriteBlankLines (1)
        Next 'Object Permission
      End If ' user defined role
      sTextStream.Close
      DoIt = 0
    End If 'DoIt
  Next 'oRole

End Function


Function ScriptRules()

  Dim oRule

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_ObjectPermissions _
                Or SQLDMOScript_Default _
                Or SQLDMOScript_Bindings
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Rules\"

  For Each oRule In oDatabase.defaults
    sFileName = sWorkingFolder & oRule.Name & ".RUL"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write "use " & oDatabase.Name
    sTextStream.WriteBlankLines (1)
    sTextStream.Write "GO"
    sTextStream.WriteBlankLines (1)
    sTextStream.Write oRule.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'oRule

End Function


Function ScriptTables()

  Dim oTable

  ' never script drop table
  iScriptType = SQLDMOScript_ObjectPermissions _
                Or SQLDMOScript_OwnerQualify _
                Or SQLDMOScript_Default _
                Or SQLDMOScript_Indexes _
                Or SQLDMOScript_DRI_All
  iScript2Type = SQLDMOScript2_NoWhatIfIndexes
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Tables\"

  For Each oTable In oDatabase.Tables
    If Not oTable.SystemObject Then
      sFileName = sWorkingFolder & oTable.owner & "." & oTable.Name & ".TAB"
      oFileSystem.CreateTextFile sFileName, True
      Set oFile = oFileSystem.GetFile(sFileName)
      Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
      sTextStream.Write "use " & oDatabase.Name
      sTextStream.WriteBlankLines (1)
      sTextStream.Write "GO"
      sTextStream.WriteBlankLines (1)
      sTextStream.Write oTable.Script(iScriptType, , , iScript2Type)
      sTextStream.Close
    End If
  Next 'oTable

End Function


Function ScriptTriggers()

  Dim oTable, oTrigger

  ' if you put the var in the createfolder you lose the trailing whack
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Triggers\"
  
  For Each oTable In oDatabase.Tables
    For Each oTrigger In oTable.Triggers
      If Not oTable.SystemObject Then

        iScriptType = SQLDMOScript_Drops _
                      Or SQLDMOScript_OwnerQualify _
                      Or SQLDMOScript_Default
        iScript2Type = SQLDMOScript2_Default
        sFileName = sWorkingFolder & oTrigger.owner & "." & oTrigger.Name & ".TRG"

        oFileSystem.CreateTextFile sFileName, True
        Set oFile = oFileSystem.GetFile(sFileName)
        Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
        sTextStream.Write "use " & oDatabase.Name
        sTextStream.WriteBlankLines (1)
        sTextStream.Write "GO"
        sTextStream.WriteBlankLines (1)
        sTextStream.Write oTrigger.Script(iScriptType, , iScript2Type)
        sTextStream.Close
      End If
    Next 'oTrigger
  Next 'oTable

End Function


Function ScriptUsers()

  Dim oUser

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_Permissions _
                Or SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default _
                 Or SQLDMOScript2_LoginSID
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Users\"

  For Each oUser In oDatabase.Users
    sFileName = sWorkingFolder & GoodFileName(oUser.Name) & ".USR"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write "use " & oDatabase.Name
    sTextStream.WriteBlankLines (1)
    sTextStream.Write "GO"
    sTextStream.WriteBlankLines (1)
    If oUser.Login <> "" Then
      sTextStream.WriteBlankLines (2)
      If oUser.Login <> "sa" Then ' never drop sa
        sTextStream.Write oSQLServer.Logins(oUser.Login).Script(iScriptType, , iScript2Type)
      Else
        sTextStream.Write oSQLServer.Logins(oUser.Login).Script(SQLDMOScript_Default, , iScript2Type)
      End If
    End If
    sTextStream.WriteBlankLines (2)
    If oUser.Name <> "dbo" Then 'never drop dbo
      sTextStream.Write oUser.Script(iScriptType, , iScript2Type)
    Else
      sTextStream.Write oUser.Script(SQLDMOScript_Permissions Or SQLDMOScript_Default, , iScript2Type)
    End If
    sTextStream.Close
  Next 'oRule

End Function


Function ScriptViews()

  Dim oView

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_ObjectPermissions _
                Or SQLDMOScript_OwnerQualify _
                Or SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default
  sWorkingFolder = sWorkingFolderRoot & oDatabase.Name & "\Views\"

  For Each oView In oDatabase.Views
    If Not oView.SystemObject Then
      sFileName = sWorkingFolder & oView.owner & "." & oView.Name & ".PRC"
      oFileSystem.CreateTextFile sFileName, True
      Set oFile = oFileSystem.GetFile(sFileName)
      Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
      sTextStream.Write "use " & oDatabase.Name
      sTextStream.WriteBlankLines (1)
      sTextStream.Write "GO"
      sTextStream.WriteBlankLines (1)
      sTextStream.Write oView.Script(iScriptType, , iScript2Type)
      sTextStream.Close
    End If
  Next 'oView

End Function


Function ScriptAlerts()

  Dim oAlert
  
  iScriptType = SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default

  For Each oAlert In oSQLServer.JobServer.Alerts
    sFileName = sWorkingFolderRoot & "\AgentAlerts\" & GoodFileName(oAlert.Name) & ".ALR"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write oAlert.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'Alert

End Function


Function ScriptBackupDevices()

  Dim oBackupDevice

  iScriptType = SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default

  For Each oBackupDevice In oSQLServer.BackupDevices
    sFileName = sWorkingFolderRoot & "\BackupDevices\" & GoodFileName(oBackupDevice.Name) & ".BDV"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write oBackupDevice.Script(iScriptType, , , iScript2Type)
    sTextStream.Close
  Next 'BackupDevice

End Function


Function ScriptJobs()

  Dim oJob

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_Default _
                Or SQLDMOScript_IncludeIfNotExists _
                Or SQLDMOScript_OwnerQualify
  iScript2Type = SQLDMOScript2_AgentNotify _
                 Or SQLDMOScript2_AgentAlertJob

  For Each oJob In oSQLServer.JobServer.Jobs
    sFileName = sWorkingFolderRoot & "\Jobs\" & GoodFileName(oJob.Name) & ".JOB"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write oJob.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'Job

End Function


Function ScriptLogins()

  Dim oLogin

  iScriptType = SQLDMOScript_Drops _
                Or SQLDMOScript_Default _
                Or SQLDMOScript_IncludeIfNotExists
  iScript2Type = SQLDMOScript2_LoginSID _
                 Or SQLDMOScript2_Default

  For Each oLogin In oSQLServer.Logins
    sFileName = sWorkingFolderRoot & "\Logins\" & GoodFileName(oLogin.Name) & ".LGN"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write oLogin.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'Login

End Function


Function ScriptOperators()

  Dim oOperator
  
  iScriptType = SQLDMOScript_Default
  iScript2Type = SQLDMOScript2_Default

  For Each oOperator In oSQLServer.JobServer.Operators
    sFileName = sWorkingFolderRoot & "\AgentOperators\" & GoodFileName(oOperator.Name) & ".OPR"
    oFileSystem.CreateTextFile sFileName, True
    Set oFile = oFileSystem.GetFile(sFileName)
    Set sTextStream = oFile.OpenAsTextStream(ForWriting, TristateUseDefault)
    sTextStream.Write oOperator.Script(iScriptType, , iScript2Type)
    sTextStream.Close
  Next 'Operator

End Function


Function GoodFileName(sObjectName)

  sObjectName = Replace(sObjectName, "\", "~")
  sObjectName = Replace(sObjectName, "/", "~")
  sObjectName = Replace(sObjectName, ":", "~")
  sObjectName = Replace(sObjectName, "*", "~")
  sObjectName = Replace(sObjectName, "?", "~")
  sObjectName = Replace(sObjectName, """", "~")
  sObjectName = Replace(sObjectName, "<", "~")
  sObjectName = Replace(sObjectName, ">", "~")
  sObjectName = Replace(sObjectName, "|", "~")

  GoodFileName = sObjectName

End Function


