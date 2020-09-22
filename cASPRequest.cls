VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cASPRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Form As New Dictionary
Public Querystring As New Dictionary
Public Cookies As New Dictionary
Public Headers As New Dictionary

Public Sub ClearAll()
    Form.RemoveAll
    Querystring.RemoveAll
    Cookies.RemoveAll
    Headers.RemoveAll
End Sub

Public Sub AddValues(sTMP As String, sDic As Dictionary, MainSep As String, SubSep As String)
On Error GoTo EH
Dim SParam() As String
Dim sValue() As String
Dim I As Long

sDic.RemoveAll

SParam = Split(sTMP, MainSep)

If UBound(SParam) <> -1 Then
    For I = LBound(SParam) To UBound(SParam)
    'Param SubSep Value MainSep
        sValue = Split(Trim(SParam(I)), SubSep)
        If UBound(sValue) = 1 Then
            sDic.Add Trim(sValue(0)), Trim(sValue(1))
        End If
    Next I
End If


Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Adding Headers"
Exit Sub
End Sub

