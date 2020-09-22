VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cASPResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'To Do
'Add Response Variables

Private mvarReturnText As String
Private mvarError As String

Public Sub StartResponse()
mvarError = ""
mvarReturnText = ""
End Sub

Public Property Let GetLastError(ByVal vData As String)
    mvarError = vData
End Property

Public Property Get GetLastError() As String
    GetLastError = mvarError
End Property

Public Property Let ReturnText(ByVal vData As String)
    mvarReturnText = vData
End Property

Public Property Get ReturnText() As String
    ReturnText = mvarReturnText
End Property

Public Sub sWrite(sTXT As String)
'Response.Write
mvarReturnText = mvarReturnText + sTXT
End Sub



