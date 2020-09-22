VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScripter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mvarScriptCode As String

Public Sub AddCode(sTXT As String)
    mvarScriptCode = mvarScriptCode & sTXT
End Sub

Public Property Let ScriptCode(ByVal vData As String)
    mvarScriptCode = vData
End Property

Public Property Get ScriptCode() As String
    ScriptCode = mvarScriptCode
End Property


