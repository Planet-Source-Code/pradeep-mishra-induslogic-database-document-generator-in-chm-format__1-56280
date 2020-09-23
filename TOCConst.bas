Attribute VB_Name = "TOCConst"
Option Explicit

Public TOC As String

Public Const TOCHeader = "<!DOCTYPE HTML PUBLIC '-//IETF//DTD HTML//EN'>" & _
    "<HTML>" & _
    "<HEAD>" & _
    "<meta name='GENERATOR' content='Microsoft&reg; HTML Help Workshop 4.1'>" & _
    "<!-- Sitemap 1.0 -->" & _
    "</HEAD><BODY>" & _
    "<OBJECT type='text/site properties'>" & _
    "<param name='ImageType' value='Folder'>" & _
    "</OBJECT><UL>" & vbCrLf

Public Const TOCServer = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%SERVER%>'>" & _
    "<param name='Local' value=''>" & _
    "<param name='ImageNumber' value='1'>" & _
    "</OBJECT><UL>" & vbCrLf

Public Const TOCDatabase = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%DATABASE%>'>" & _
    "<param name='Local' value='db_details.htm'>" & _
    "<param name='ImageNumber' value='1'>" & _
    "</OBJECT><UL>" & vbCrLf
Public Const TOCDatabaseD = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='Database Details'>" & _
    "<param name='Local' value='db_details.htm'>" & _
    "</OBJECT>" & vbCrLf

Public Const TOCTableH = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='Tables'>" & _
    "<param name='Local' value='<%TABLELINK%>'>" & _
    "</OBJECT><UL>" & vbCrLf
Public Const TOCTableD = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%TABLENAME%>'>" & _
    "<param name='Local' value='<%TABLELINK%>'>" & _
    "</OBJECT>" & vbCrLf

Public Const TOCTriggerH = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='Triggers'>" & _
    "<param name='Local' value=''>" & _
    "</OBJECT><UL>" & vbCrLf
Public Const TOCTriggerD = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%TRIGGERNAME%>'>" & _
    "<param name='Local' value='<%TRIGGERLINK%>'>" & _
    "</OBJECT>" & vbCrLf

Public Const TOCViewH = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='Views'>" & _
    "<param name='Local' value=''>" & _
    "</OBJECT><UL>" & vbCrLf
Public Const TOCViewD = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%VIEWNAME%>'>" & _
    "<param name='Local' value='<%VIEWLINK%>'>" & _
    "</OBJECT>" & vbCrLf


Public Const TOCStoredProcedureH = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='Stored Procedures'>" & _
    "<param name='Local' value=''>" & _
    "</OBJECT><UL>" & vbCrLf
Public Const TOCStoredProcedureD = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%STOREDPROCEDURENAME%>'>" & _
    "<param name='Local' value='<%STOREDPROCEDURELINK%>'>" & _
    "</OBJECT>" & vbCrLf

Public Const TOCUserDefinedFunctionH = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='User Defined Functions'>" & _
    "<param name='Local' value=''>" & _
    "</OBJECT><UL>" & vbCrLf
Public Const TOCUserDefinedFunctionD = "<LI> <OBJECT type='text/sitemap'>" & _
    "<param name='Name' value='<%USERDEFINEDFUNCTIONNAME%>'>" & _
    "<param name='Local' value='<%USERDEFINEDFUNCTIONLINK%>'>" & _
    "</OBJECT>" & vbCrLf

