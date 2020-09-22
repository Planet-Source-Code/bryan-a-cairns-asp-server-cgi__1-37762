VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScriptControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Script control class

Public WithEvents cScript As ScriptControl
Attribute cScript.VB_VarHelpID = -1
Private m_Response As cASPResponse
Private m_Request As cASPRequest
Private m_Server As cASPServer

Public Sub StartScript(Response As cASPResponse, Request As cASPRequest, Server As cASPServer, sCode As String)
On Error Resume Next
'Set cScript = Form1.ScriptControl1
'Response.sWrite cScrClass.ScriptCode
'Set our local variable
Set m_Response = Response
Set m_Request = Request
Set m_Server = Server

cScript.AddObject "Response", m_Response, True
cScript.AddObject "Request", m_Request, True
cScript.AddObject "Server", m_Server, True

'add all variables that are declared in the script
'cScript.AddObject "rs", CreateObject("ADODB.recordset"), True
'cScript.AddObject "conn", CreateObject("ADODB.Connection"), True

'add the code and execute it
cScript.ExecuteStatement sCode
End Sub


Public Sub AddCode(sCode As String)
    cScript.AddCode sCode
End Sub

Public Sub AddObject(Name As String, Object As Object, Optional AddMembers As Boolean)
    cScript.AddObject Name, Object, AddMembers
End Sub

Public Sub Eval(Expression As String)
    cScript.Eval Expression
End Sub

Public Sub ExecuteStatement(Statment As String)
    cScript.ExecuteStatement Statment
End Sub

Private Sub cScript_Error()
'an error occured
m_Response.sWrite "<BR><BR>ASP Error: #" & cScript.Error.Number & " " & cScript.Error.Description & "<BR>"
m_Response.sWrite "Line: " & cScript.Error.Line & " Col:" & cScript.Error.Column & "<BR>" & "<BR>"
End Sub

Private Sub cScript_Timeout()
'the script timed out
m_Response.sWrite "<BR><BR>ASP Error: Script Timeout<BR><BR>"
End Sub
