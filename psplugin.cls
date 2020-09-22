VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "psplugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Power Server Plugin Template
'These are the default subs and functions called by the server
'if you remove one an error will occure
'!!!The server is multi-threaded so do not use global variables!!!

Function Plugin_Name() As String
'Called when the server needs to know this files friendly name
Plugin_Name = "ASP"
End Function

Function Plugin_Effect() As String
'Called when the server needs to know what file types this plugin effects
'You must return a comma seperated string of file extentions
'example: Plugin_Effect = ".html,.htm,.asp,.shtml"
'To effect all files use a ".*"
'To effect none but still get notification us "none" - please note this might interfere with script languages!
Plugin_Effect = ".asp,.inc"
End Function

Public Sub Server_Start()
'Called when the server is started

End Sub

Public Sub Server_Stop()
'Called when the server is stopped

End Sub

Function Command_Get(sCommand As String, sParams As String, sHeaders As String, sIP As String, sLocalPath As String, sWebPath As String) As String
'Called when a user requests a file from the server
'You must return a filename or the raw data of the file
'The server will check to see if it is a filename by checking for the existance of ":" as the second char
'example: c:\test.txt is a filename c\test.txt would not be a filename
'Also the server checks to see if the file exists, if not then the Command_Get is treated as raw data.

'set the root path
'LocalRoot = ""

Command_Get = ServASPPage(sCommand, sParams, sHeaders, sIP, sLocalPath, sWebPath)
End Function

Public Sub On_ActiveState(bAllow As Boolean)
'Called when the server changes the active state of this plugin

End Sub

Public Sub On_Status(sStatusText As String)
'Called when the server status changes

End Sub

Public Sub On_Blacklist(sIP As String)
'Called when the server encounters a blacklisted IP address

End Sub

Public Sub ShowConfig()
'Called when the user wants to configure this plugin
Dim FRM As frmconfig
Set FRM = New frmconfig
FRM.Show vbModal
End Sub

Public Sub ShowAbout()
'Called when the user wants to see the plugins about window
Dim FRM As frmabout
Set FRM = New frmabout
FRM.Show vbModal
End Sub
